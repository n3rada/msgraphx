# msgraphx/modules/sharepoint/sites.py
#
# Enumerate SharePoint sites accessible to the current user.
#
# Approach: POST /search/query with EntityType.Site
# --------------------------------------------------
# The Graph Search API with EntityType.Site returns every site the calling
# user can reach, regardless of *how* they got access:
#
#   - Via M365 Unified group membership (direct or nested/transitive)
#   - Via direct sharing (a site shared with the user individually)
#   - Via link sharing (a site URL shared with them)
#   - Sites the user is explicitly following
#
# This is a single API call and is more complete than the previous approach,
# which enumerated group membership then resolved each group's site one by one
# (N+1 calls, missing directly shared sites entirely).
#
# For app-only tokens the search returns all tenant sites — useful for
# organisation-wide recon.
#
# Enrichment (--enrich, delegated only)
# --------------------------------------
# For each M365 Unified group the user belongs to, resolve the group's root
# SharePoint site via GET /groups/{id}/sites/root (parallelised). Sites that
# match a group get access_via="group"; the rest get access_via="direct".
# This tells the operator *how* they have access without re-running a separate
# groups command.

from __future__ import annotations

import argparse
import asyncio

from loguru import logger
from msgraph.generated.models.entity_type import EntityType
from msgraph.generated.models.site import Site
from rich.table import Table

from ...core import graph_search
from ...core.context import GraphContext
from ...utils import cache, output
from ...utils.console import console
from ...utils.errors import handle_graph_errors
from ...utils.roles import require_scopes
from .groups import get_user_m365_groups

async def fetch(context: GraphContext) -> list[dict]:
    """Return SharePoint sites accessible to the current user.

    Uses POST /search/query (EntityType.Site) — one call that covers all
    access paths: group membership, direct sharing, followed sites, and
    public sites the user has visited.

    For app-only tokens this returns all sites in the tenant.
    """
    options = graph_search.SearchOptions(
        query_string="*",
        # Site search does not support sorting by createdDateTime in all
        # tenants; disable to avoid API errors.
        sort_by=None,
        page_size=500,
        region=context.region if context.is_app_only else None,
    )

    sites: list[dict] = []
    async for raw in graph_search.search_entities(
        context.graph_client,
        entity_types=[EntityType.Site],
        options=options,
    ):
        if not isinstance(raw, Site):
            continue

        sites.append({
            "id": raw.id,
            "name": raw.display_name or raw.name or "",
            "description": raw.description or "",
            "web_url": raw.web_url or "",
            "created_datetime": (
                raw.created_date_time.isoformat() if raw.created_date_time else None
            ),
            "last_modified_datetime": (
                raw.last_modified_date_time.isoformat()
                if raw.last_modified_date_time
                else None
            ),
        })

    return sites

def _site_obj_to_dict(site, group_info: dict | None = None) -> dict:
    """Convert a Site SDK object to a plain dict, optionally embedding group info."""
    d = {
        "id": site.id,
        "name": site.display_name or site.name or "",
        "description": site.description or "",
        "web_url": site.web_url or "",
        "created_datetime": (
            site.created_date_time.isoformat() if site.created_date_time else None
        ),
        "last_modified_datetime": (
            site.last_modified_date_time.isoformat() if site.last_modified_date_time else None
        ),
    }
    if group_info:
        d.update(group_info)
    return d

async def _resolve_group_sites(context: GraphContext, group) -> list[tuple[str, dict]]:
    """Return all (site_id, group_info) pairs for every site owned by a group.

    Uses GET /groups/{id}/sites (collection) rather than /sites/root so that
    groups with multiple associated site collections are fully covered.
    """
    info = {
        "group_id": group.id,
        "group_name": group.display_name or "",
        "group_visibility": group.visibility or "",
    }
    results = []
    try:
        site = await context.graph_client.groups.by_group_id(group.id).sites.by_site_id("root").get()
        if site and site.id:
            results.append((site.id, info))
    except Exception:
        pass
    return results

async def _fetch_group_sites(context: GraphContext, group) -> list[dict]:
    """Return full site dicts for every site owned by a group, with group info embedded."""
    info = {
        "group_id": group.id,
        "group_name": group.display_name or "",
        "group_visibility": group.visibility or "",
    }
    results = []
    try:
        site = await context.graph_client.groups.by_group_id(group.id).sites.by_site_id("root").get()
        if site and site.id:
            results.append(_site_obj_to_dict(site, info))
    except Exception:
        pass
    return results

async def fetch_from_groups(context: GraphContext) -> list[dict]:
    """Fetch SharePoint sites directly from the user's M365 group membership.

    Bypasses the tenant-wide site search entirely: one call to get groups,
    then one call per group to get its sites (parallelised). Returns only
    sites the user has access to via group membership, with group info embedded.
    """
    groups = await get_user_m365_groups(context.graph_client)
    if not groups:
        return []

    logger.info(f"Fetching sites for {len(groups)} M365 group(s)...")
    nested = await asyncio.gather(*[_fetch_group_sites(context, g) for g in groups])

    seen: set[str] = set()
    sites: list[dict] = []
    for group_sites in nested:
        for site in group_sites:
            if site["id"] not in seen:
                seen.add(site["id"])
                sites.append(site)

    return sites

async def enrich_with_groups(context: GraphContext, sites: list[dict]) -> list[dict]:
    """Annotate each site with the M365 group it belongs to, if any.

    Resolves each of the user's M365 Unified groups to its root SharePoint
    site in parallel, then cross-references against the sites list. Sites
    that match a group get access_via='group'; others get access_via='direct'.
    """
    groups = await get_user_m365_groups(context.graph_client)
    if not groups:
        logger.warning("No M365 Unified groups found — skipping group enrichment.")
        return [{**s, "access_via": "direct", "group_id": None, "group_name": None, "group_visibility": None} for s in sites]

    logger.info(f"Resolving sites for {len(groups)} M365 group(s)...")
    nested = await asyncio.gather(*[_resolve_group_sites(context, g) for g in groups])

    group_site_map: dict[str, dict] = {}
    for pairs in nested:
        for site_id, info in pairs:
            group_site_map[site_id] = info

    logger.debug(f"{len(group_site_map)} group-owned site(s) resolved")

    enriched = []
    for site in sites:
        match = group_site_map.get(site["id"])
        if match:
            enriched.append({
                **site,
                "access_via": "group",
                "group_id": match["group_id"],
                "group_name": match["group_name"],
                "group_visibility": match["group_visibility"],
            })
        else:
            enriched.append({
                **site,
                "access_via": "direct",
                "group_id": None,
                "group_name": None,
                "group_visibility": None,
            })

    matched = sum(1 for s in enriched if s["access_via"] == "group")
    logger.info(f"{matched} site(s) matched to a group, {len(enriched) - matched} via direct access")
    return enriched

def add_arguments(parser: argparse.ArgumentParser) -> None:
    parser.add_argument(
        "--from-groups",
        action="store_true",
        dest="from_groups",
        help=(
            "Fetch sites directly from your M365 group membership — skips the tenant-wide "
            "site search entirely. Fast: one API call per group (parallelised). "
            "Returns only sites backed by a group you belong to (direct or transitive). "
            "Requires delegated auth."
        ),
    )
    parser.add_argument(
        "--enrich",
        action="store_true",
        help=(
            "Cross-reference all accessible sites against your M365 group membership, "
            "tagging each site as 'group' or 'direct'. Use --from-groups instead when "
            "you only care about group-backed sites. Requires delegated auth."
        ),
    )

@handle_graph_errors
@require_scopes("Sites.Read.All")
async def run_with_arguments(context: GraphContext, args: argparse.Namespace) -> int:
    from_groups = getattr(args, "from_groups", False)
    enrich = getattr(args, "enrich", False)

    if from_groups and enrich:
        logger.error("--from-groups and --enrich are mutually exclusive.")
        return 1

    if from_groups:
        if context.is_app_only:
            logger.error("--from-groups requires delegated authentication (user context).")
            return 1
        sites = await fetch_from_groups(context)
        if not sites:
            logger.info("No group-backed sites found.")
            if context.json_output:
                output.print_json([])
            return 0
        cache.save_results(sites, key="sites_groups", identity=context.identity_hash)
        if context.json_output:
            output.print_json(sites)
            return 0
        if context.ndjson_output:
            for site in sites:
                output.print_ndjson_item(site)
            return 0
        _print_sites(sites, show_group=True)
        logger.success(f"{len(sites)} group-backed site(s) found")
        return 0

    logger.info("Enumerating accessible SharePoint sites via search index")
    sites = await fetch(context)

    if not sites:
        logger.info("No sites found.")
        if context.json_output:
            output.print_json([])
        return 0

    if enrich:
        if context.is_app_only:
            logger.error("--enrich requires delegated authentication (user context).")
            return 1
        cached = cache.load_results(key="sites", identity=context.identity_hash)
        if cached:
            logger.info(f"Using {len(cached)} cached site(s) — skipping search API. Run 'sp sites' first to refresh.")
            sites = cached
        else:
            cache.save_results(sites, key="sites", identity=context.identity_hash)
        sites = await enrich_with_groups(context, sites)
        cache.save_results(sites, key="sites_enriched", identity=context.identity_hash)
    else:
        cache.save_results(sites, key="sites", identity=context.identity_hash)

    if context.json_output:
        output.print_json(sites)
        return 0

    if context.ndjson_output:
        for site in sites:
            output.print_ndjson_item(site)
        return 0

    _print_sites(sites, show_group=enrich)
    logger.success(f"{len(sites)} site(s) found")
    return 0

def _print_sites(sites: list[dict], show_group: bool = False) -> None:
    if not sites:
        return

    table = Table(
        title=f"SharePoint sites ({len(sites)})",
        show_header=True,
        header_style="bold",
        box=None,
        padding=(0, 1),
    )
    table.add_column("#", justify="right", style="dim")
    table.add_column("Site")
    table.add_column("URL")
    if show_group:
        table.add_column("Group")

    for idx, site in enumerate(sites, 1):
        if show_group:
            group_name = site.get("group_name") or ""
            visibility = site.get("group_visibility") or ""
            access_via = site.get("access_via", "group")
            if group_name:
                vis_tag = f" [dim]({visibility})[/dim]" if visibility else ""
                group_cell = f"[cyan]{group_name}[/cyan]{vis_tag}"
            else:
                group_cell = "[dim]direct[/dim]" if access_via == "direct" else ""
            table.add_row(str(idx), site["name"] or "(unnamed)", site["web_url"], group_cell)
        else:
            table.add_row(str(idx), site["name"] or "(unnamed)", site["web_url"])

    console.print(table)
