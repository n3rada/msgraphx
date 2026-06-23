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
#
# Required delegated permissions: Sites.Read.All
# Required application permissions: Sites.Read.All + region set

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


async def _resolve_group_site(context: GraphContext, group) -> tuple[str, dict] | None:
    """Return (site_id, group_info) for a group's root SharePoint site, or None on failure."""
    try:
        site = await context.graph_client.groups.by_group_id(group.id).sites.by_site_id("root").get()
        if site and site.id:
            return site.id, {
                "group_id": group.id,
                "group_name": group.display_name or "",
                "group_visibility": group.visibility or "",
            }
    except Exception:
        pass
    return None


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

    logger.info(f"Resolving root sites for {len(groups)} M365 group(s)...")
    results = await asyncio.gather(*[_resolve_group_site(context, g) for g in groups])

    group_site_map: dict[str, dict] = {}
    for result in results:
        if result:
            site_id, info = result
            group_site_map[site_id] = info

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
        "--enrich",
        action="store_true",
        help=(
            "Resolve how you have access to each site: cross-references your M365 Unified "
            "group membership and tags each site as 'group' (backed by a group you belong to) "
            "or 'direct' (shared directly or via link). Requires delegated auth. "
            "Makes one extra API call per group."
        ),
    )


@handle_graph_errors
async def run_with_arguments(context: GraphContext, args: argparse.Namespace) -> int:
    logger.info("Enumerating accessible SharePoint sites via search index")

    sites = await fetch(context)

    if not sites:
        logger.info("No sites found.")
        if context.json_output:
            output.print_json([])
        return 0

    enrich = getattr(args, "enrich", False)

    if enrich:
        if context.is_app_only:
            logger.error("--enrich requires delegated authentication (user context).")
            return 1
        sites = await enrich_with_groups(context, sites)

    cache.save_results(sites, key="sites", identity=context.identity_hash)

    if context.json_output:
        output.print_json(sites)
        return 0

    if context.ndjson_output:
        for site in sites:
            output.print_ndjson_item(site)
        return 0

    table = Table(
        title=f"Accessible SharePoint sites ({len(sites)})",
        show_header=True,
        header_style="bold",
        box=None,
        padding=(0, 1),
    )
    table.add_column("#", justify="right", style="dim")
    table.add_column("Site")
    table.add_column("URL")
    if enrich:
        table.add_column("Access")

    for idx, site in enumerate(sites, 1):
        if enrich:
            if site["access_via"] == "group":
                visibility = site["group_visibility"]
                vis_tag = f" [dim]({visibility})[/dim]" if visibility else ""
                access_cell = f"[cyan]{site['group_name']}[/cyan]{vis_tag}"
            else:
                access_cell = "[dim]direct[/dim]"
            table.add_row(str(idx), site["name"] or "(unnamed)", site["web_url"], access_cell)
        else:
            table.add_row(str(idx), site["name"] or "(unnamed)", site["web_url"])

    console.print(table)
    logger.success(f"{len(sites)} site(s) found")
    return 0
