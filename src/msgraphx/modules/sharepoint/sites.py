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
# Required delegated permissions: Sites.Read.All
# Required application permissions: Sites.Read.All + region set

from __future__ import annotations

import argparse

from loguru import logger
from msgraph.generated.models.entity_type import EntityType
from msgraph.generated.models.site import Site
from rich.table import Table

from ...core import graph_search
from ...core.context import GraphContext
from ...utils import cache, output
from ...utils.console import console
from ...utils.errors import handle_graph_errors


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


def add_arguments(parser: argparse.ArgumentParser) -> None:
    pass


@handle_graph_errors
async def run_with_arguments(context: GraphContext, args: argparse.Namespace) -> int:
    logger.info("Enumerating accessible SharePoint sites via search index")

    sites = await fetch(context)

    if not sites:
        logger.info("No sites found.")
        if context.json_output:
            output.print_json([])
        return 0

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

    for idx, site in enumerate(sites, 1):
        table.add_row(str(idx), site["name"] or "(unnamed)", site["web_url"])

    console.print(table)
    logger.success(f"{len(sites)} site(s) found")
    return 0
