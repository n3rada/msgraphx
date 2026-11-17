# msgraphx/modules/me/trending.py
#
# Show documents trending around the current user.
# Uses GET /me/insights/trending which surfaces files from OneDrive and
# SharePoint that are relevant based on the user's closest network activity.

# Built-in imports
from __future__ import annotations

import argparse
from typing import Any

# External library imports
from kiota_abstractions.base_request_configuration import RequestConfiguration
from loguru import logger
from msgraph.generated.models.trending import Trending
from msgraph.generated.users.item.insights.trending.trending_request_builder import (
    TrendingRequestBuilder,
)
from rich.table import Table

# Local library imports
from ...core.context import GraphContext
from ...utils import output, pagination
from ...utils.console import console
from ...utils.errors import handle_graph_errors
from ...utils.roles import require_scopes

_MAX_PAGE_SIZE = 100

async def fetch(
    context: GraphContext,
    top: int = 25,
    type_filter: str | None = None,
) -> list[dict[str, Any]]:
    """Return documents trending around the current user as plain dicts.

    The type filter is applied client-side. Pages are followed until `top`
    matching items are collected or the results run out.

    Args:
        context: Shared Graph runtime context (delegated auth).
        top: Maximum number of matching items to return. Each page requests
            at most 100.
        type_filter: Case-insensitive substring matched against the
            resource type (e.g. `Excel`). `None` keeps all items.

    Returns:
        One dict per item with the keys `rank` (1-based, among returned
        items), `type`, `title`, `site`, `weight`, `url`, and `id`.

    Raises:
        AuthenticationError: If the Graph token is invalid or expired.
        ODataError: If the Graph API rejects the request.
    """
    query_params = TrendingRequestBuilder.TrendingRequestBuilderGetQueryParameters(
        top=min(top, _MAX_PAGE_SIZE),
    )
    config = RequestConfiguration(query_parameters=query_params)

    low = type_filter.lower() if type_filter else None
    items: list[Trending] = []
    item: Trending
    async for item in pagination.GraphPaginator(
        context.graph_client.me.insights.trending, config
    ):
        viz = item.resource_visualization
        if low and not (viz and viz.type and low in viz.type.lower()):
            continue
        items.append(item)
        if len(items) >= top:
            break

    return [
        {
            "rank": i,
            "type": (item.resource_visualization.type if item.resource_visualization else None),
            "title": (item.resource_visualization.title if item.resource_visualization else None),
            "site": (item.resource_visualization.container_display_name if item.resource_visualization else None),
            "weight": item.weight,
            "url": (item.resource_reference.web_url if item.resource_reference else None),
            "id": (item.resource_reference.id if item.resource_reference else None),
        }
        for i, item in enumerate(items, 1)
    ]

def add_arguments(parser: argparse.ArgumentParser) -> None:
    """Register the `me trending` options on the parser.

    Args:
        parser: Parser for the `me trending` subcommand.
    """
    parser.add_argument(
        "--top",
        "-n",
        type=int,
        default=25,
        metavar="N",
        help="Number of trending items to display (default: 25).",
    )

    parser.add_argument(
        "--type",
        "-t",
        type=str,
        default=None,
        metavar="TYPE",
        help="Filter by file type (e.g. Excel, PowerPoint, Word, Pdf).",
    )

@handle_graph_errors
@require_scopes("Sites.Read.All")
async def run_with_arguments(
    context: GraphContext, args: argparse.Namespace
) -> int:
    """List documents trending around the current user.

    Prints a table, or JSON/NDJSON when the matching output flag is set.

    Args:
        context: Shared Graph runtime context.
        args: Parsed CLI arguments (`top`, `type`).

    Returns:
        0 on success (including no items), 1 for app-only auth.

    Raises:
        AuthenticationError: If the Graph token is invalid or expired.
        ODataError: If the Graph API rejects the request.
    """
    if context.is_app_only:
        logger.error("This module requires delegated authentication (user context).")
        return 1

    logger.info("Fetching trending documents")

    rows = await fetch(context, top=args.top, type_filter=args.type)

    if not rows:
        logger.info("No trending items found.")
        if context.json_output:
            output.print_json([])
        return 0

    if context.json_output:
        output.print_json(rows)
        return 0

    if context.ndjson_output:
        for row in rows:
            output.print_ndjson_item(row)
        return 0

    table = Table(show_header=True, header_style="bold", box=None, padding=(0, 1))
    table.add_column("#", style="dim", justify="right", width=4)
    table.add_column("Type", style="cyan", width=12)
    table.add_column("Title", min_width=40)
    table.add_column("Site", style="dim", max_width=30)
    table.add_column("Weight", style="yellow", justify="right", width=8)

    for row in rows:
        title = row["title"] or "(untitled)"
        if len(title) > 80:
            title = title[:77] + "..."
        weight = f"{row['weight']:.3f}" if row["weight"] else ""
        table.add_row(str(row["rank"]), row["type"] or "?", title, row["site"] or "", weight)

    console.print("[bold]Trending around you[/bold]")
    console.rule()
    console.print(table)
    logger.success(f"{len(rows)} trending item(s).")

    return 0
