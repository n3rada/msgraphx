# msgraphx/modules/me/trending.py
#
# Show documents trending around the current user.
# Uses GET /me/insights/trending which surfaces files from OneDrive and
# SharePoint that are relevant based on the user's closest network activity.
#
# Required delegated permission: Sites.Read.All

# Built-in imports
from __future__ import annotations

import argparse

# External library imports
from kiota_abstractions.base_request_configuration import RequestConfiguration
from loguru import logger
from msgraph.generated.users.item.insights.trending.trending_request_builder import (
    TrendingRequestBuilder,
)
from rich.table import Table

# Local library imports
from ...core.context import GraphContext
from ...utils import output
from ...utils.console import console
from ...utils.errors import handle_graph_errors


async def fetch(
    context: GraphContext,
    top: int = 25,
    type_filter: str | None = None,
) -> list[dict]:
    """Return trending documents for the current user as plain dicts.

    Raises on API error — callers are responsible for handling exceptions.
    """
    query_params = TrendingRequestBuilder.TrendingRequestBuilderGetQueryParameters(
        top=min(top, 100),
    )
    config = RequestConfiguration(query_parameters=query_params)

    result = await context.graph_client.me.insights.trending.get(
        request_configuration=config
    )
    items = result.value if result and result.value else []

    if type_filter:
        low = type_filter.lower()
        items = [
            i for i in items
            if i.resource_visualization
            and i.resource_visualization.type
            and low in i.resource_visualization.type.lower()
        ]

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


def add_arguments(parser: "argparse.ArgumentParser") -> None:
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
async def run_with_arguments(
    context: "GraphContext", args: "argparse.Namespace"
) -> int:
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
