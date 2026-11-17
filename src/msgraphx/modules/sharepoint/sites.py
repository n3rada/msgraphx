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
from collections.abc import Mapping, Sequence
from typing import Any

from loguru import logger
from msgraph.generated.models.entity_type import EntityType
from msgraph.generated.models.group import Group
from msgraph.generated.models.o_data_errors.o_data_error import ODataError
from msgraph.generated.models.site import Site
from rich.table import Table

from ...core import graph_search
from ...core.context import GraphContext
from ...utils import cache, output
from ...utils.console import console
from ...utils.errors import handle_graph_errors, is_graph_auth_error
from ...utils.roles import require_scopes
from .groups import get_user_m365_groups

async def fetch(context: GraphContext) -> list[dict[str, Any]]:
    """Return the SharePoint sites accessible to the caller.

    Uses the Graph Search API (`EntityType.Site`), which covers every access
    path: group membership, direct sharing, followed sites, and public sites
    the user has visited. App-only tokens get every site in the tenant.

    Args:
        context: Shared Graph runtime context.

    Returns:
        One row per site with `id`, `name`, `description`, `web_url`,
        `created_datetime`, and `last_modified_datetime` (ISO 8601 strings
        or None).
    """
    options = graph_search.SearchOptions(
        query_string="*",
        # Site search does not support sorting by createdDateTime in all
        # tenants; disable to avoid API errors.
        sort_by=None,
        page_size=500,
        region=context.region if context.is_app_only else None,
    )

    sites: list[dict[str, Any]] = []
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

def _site_obj_to_dict(
    site: Site, group_info: Mapping[str, Any] | None = None
) -> dict[str, Any]:
    """Convert a site into an output row, optionally merging group details.

    Args:
        site: The site to convert.
        group_info: Extra keys (e.g. `group_id`, `group_name`,
            `group_visibility`) merged into the row.

    Returns:
        A row with the same keys as `fetch` returns, plus `group_info`.
    """
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

def _group_info(group: Group) -> dict[str, Any]:
    """Return the group fields embedded in site rows.

    Args:
        group: The owning M365 Unified group.

    Returns:
        A dict with `group_id`, `group_name`, and `group_visibility`.
    """
    return {
        "group_id": group.id,
        "group_name": group.display_name or "",
        "group_visibility": group.visibility or "",
    }


async def _get_group_root_site(context: GraphContext, group: Group) -> Site | None:
    """Return the group's root SharePoint site via `GET /groups/{id}/sites/root`.

    A group without a provisioned site (404) is logged at debug level; any
    other Graph error is logged as a warning. Both return None so one group
    cannot abort the whole enumeration.

    Args:
        context: Shared Graph runtime context.
        group: The M365 Unified group whose site to fetch.

    Returns:
        The root site, or None if the group has no ID, no site, or the
        request failed.

    Raises:
        ODataError: The token is invalid or expired; re-raised so the
            entry point's `handle_graph_errors` stops the command.
    """
    if not group.id:
        return None
    try:
        site = await context.graph_client.groups.by_group_id(group.id).sites.by_site_id("root").get()
    except ODataError as exc:
        if is_graph_auth_error(exc):
            raise
        label = group.display_name or group.id
        if exc.response_status_code == 404:
            logger.debug(f"Group '{label}' has no SharePoint site.")
        else:
            logger.warning(f"Failed to resolve site for group '{label}': {exc}")
        return None
    return site if site and site.id else None


async def _resolve_group_sites(
    context: GraphContext, group: Group
) -> list[tuple[str, dict[str, Any]]]:
    """Return the group's root site ID paired with the group's details.

    Args:
        context: Shared Graph runtime context.
        group: The M365 Unified group whose site to resolve.

    Returns:
        A list holding one tuple (site_id, group_info), where group_info has
        `group_id`, `group_name`, and `group_visibility`; empty if the site
        could not be resolved (see `_get_group_root_site`).

    Raises:
        ODataError: The token is invalid or expired.
    """
    site = await _get_group_root_site(context, group)
    if site is None or not site.id:
        return []
    return [(site.id, _group_info(group))]


async def _fetch_group_sites(
    context: GraphContext, group: Group
) -> list[dict[str, Any]]:
    """Return the group's root site as an output row with group details.

    Args:
        context: Shared Graph runtime context.
        group: The M365 Unified group whose site to fetch.

    Returns:
        A list holding the site row (see `_site_obj_to_dict`) with
        `group_id`, `group_name`, and `group_visibility`; empty if the site
        could not be fetched (see `_get_group_root_site`).

    Raises:
        ODataError: The token is invalid or expired.
    """
    site = await _get_group_root_site(context, group)
    if site is None:
        return []
    return [_site_obj_to_dict(site, _group_info(group))]

async def fetch_from_groups(context: GraphContext) -> list[dict[str, Any]]:
    """Return the root sites of the signed-in user's M365 Unified groups.

    Skips the tenant-wide site search: one call lists the user's groups
    (including nested membership), then each group's root site is fetched
    in parallel. Requires a delegated token.

    Args:
        context: Shared Graph runtime context.

    Returns:
        One row per distinct site, with `group_id`, `group_name`, and
        `group_visibility` of the first group that owns it.
    """
    groups = await get_user_m365_groups(context.graph_client)
    if not groups:
        return []

    logger.info(f"Fetching sites for {len(groups)} M365 group(s)...")
    nested = await asyncio.gather(*[_fetch_group_sites(context, g) for g in groups])

    seen: set[str] = set()
    sites: list[dict[str, Any]] = []
    for group_sites in nested:
        for site in group_sites:
            if site["id"] not in seen:
                seen.add(site["id"])
                sites.append(site)

    return sites

async def enrich_with_groups(
    context: GraphContext, sites: Sequence[Mapping[str, Any]]
) -> list[dict[str, Any]]:
    """Tag each site with the M365 group that owns it, if any.

    Resolves each of the signed-in user's M365 Unified groups to its root
    site in parallel, then matches those against `sites` by site ID.
    Requires a delegated token.

    Args:
        context: Shared Graph runtime context.
        sites: Site rows, each with an `id` key.

    Returns:
        Copies of `sites` with `access_via` ("group" or "direct"),
        `group_id`, `group_name`, and `group_visibility` added; the group
        fields are None for direct access.
    """
    groups = await get_user_m365_groups(context.graph_client)
    if not groups:
        logger.warning("No M365 Unified groups found — skipping group enrichment.")
        return [{**s, "access_via": "direct", "group_id": None, "group_name": None, "group_visibility": None} for s in sites]

    logger.info(f"Resolving sites for {len(groups)} M365 group(s)...")
    nested = await asyncio.gather(*[_resolve_group_sites(context, g) for g in groups])

    group_site_map: dict[str, dict[str, Any]] = {}
    for pairs in nested:
        for site_id, info in pairs:
            group_site_map[site_id] = info

    logger.debug(f"{len(group_site_map)} group-owned site(s) resolved")

    enriched: list[dict[str, Any]] = []
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
    """Register the `sharepoint sites` arguments on a parser.

    Args:
        parser: Parser for the `sites` subcommand.
    """
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
    """List accessible SharePoint sites and cache them.

    By default lists every site found through the search index (cache key
    `sites`). `--from-groups` lists only group-backed sites (`sites_groups`);
    `--enrich` tags each site as group or direct access (`sites_enriched`),
    reusing cached `sites` results when present so the search API is only
    called when there is no cache. Output is a table, JSON, or NDJSON
    depending on the output mode.

    Args:
        context: Shared Graph runtime context.
        args: Parsed arguments (`from_groups`, `enrich`).

    Returns:
        0 on success (including no sites found), 1 if both flags are given
        or either flag is used with an app-only token.

    Raises:
        AuthenticationError: The Graph token is invalid or expired.
    """
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

    if enrich and context.is_app_only:
        logger.error("--enrich requires delegated authentication (user context).")
        return 1

    cached = cache.load_results(key="sites", identity=context.identity_hash) if enrich else []
    if cached:
        logger.info(
            f"Using {len(cached)} cached site(s) — skipping search API. "
            "Run 'sharepoint sites' without --enrich to refresh."
        )
        sites = cached
    else:
        logger.info("Enumerating accessible SharePoint sites via search index")
        sites = await fetch(context)
        if sites:
            cache.save_results(sites, key="sites", identity=context.identity_hash)

    if not sites:
        logger.info("No sites found.")
        if context.json_output:
            output.print_json([])
        return 0

    if enrich:
        sites = await enrich_with_groups(context, sites)
        cache.save_results(sites, key="sites_enriched", identity=context.identity_hash)

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

def _print_sites(
    sites: Sequence[Mapping[str, Any]], show_group: bool = False
) -> None:
    """Print sites as a Rich table.

    Args:
        sites: Site rows with `name` and `web_url` keys.
        show_group: Add a Group column built from `group_name`,
            `group_visibility`, and `access_via`.
    """
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
