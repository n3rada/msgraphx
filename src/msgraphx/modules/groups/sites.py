# msgraphx/modules/groups/sites.py
#
# List SharePoint sites owned by a specific M365 Unified group.
# GET /groups/{id}/sites
#
# Required delegated permissions: Sites.Read.All, Group.Read.All
# Required application permissions: Sites.Read.All, Group.Read.All

from __future__ import annotations

import argparse

from loguru import logger
from rich.table import Table

from ...core.context import GraphContext
from ...utils import cache, output
from ...utils.console import console
from ...utils.errors import handle_graph_errors
from ...utils.roles import require_scopes


@handle_graph_errors
@require_scopes("Sites.Read.All")
async def run_with_arguments(context: GraphContext, args: argparse.Namespace) -> int:
    group_id = args.group_id
    logger.info(f"Fetching sites for group {group_id}")

    root_site = await context.graph_client.groups.by_group_id(group_id).sites.by_site_id("root").get()
    raw_sites = [root_site] if root_site else []

    rows = []
    for site in raw_sites:
        rows.append({
            "id": site.id,
            "name": site.display_name or site.name or "",
            "web_url": site.web_url or "",
            "description": site.description or "",
            "created_datetime": (
                site.created_date_time.isoformat() if site.created_date_time else None
            ),
            "last_modified_datetime": (
                site.last_modified_date_time.isoformat()
                if site.last_modified_date_time
                else None
            ),
        })

    if not rows:
        logger.info("No sites found for this group.")
        if context.json_output:
            output.print_json([])
        return 0

    cache.save_results(rows, key="group_sites", identity=context.identity_hash)

    if context.json_output:
        output.print_json(rows)
        return 0

    if context.ndjson_output:
        for row in rows:
            output.print_ndjson_item(row)
        return 0

    table = Table(
        title=f"Sites for group {group_id} ({len(rows)})",
        show_header=True,
        header_style="bold",
        box=None,
        padding=(0, 1),
    )
    table.add_column("#", justify="right", style="dim")
    table.add_column("Site")
    table.add_column("URL")

    for idx, row in enumerate(rows, 1):
        table.add_row(str(idx), row["name"] or "(unnamed)", row["web_url"])

    console.print(table)
    logger.success(f"{len(rows)} site(s) found")
    return 0


def add_arguments(parser: argparse.ArgumentParser) -> None:
    parser.add_argument(
        "group_id",
        help="Group ID (GUID) to list sites for.",
    )
