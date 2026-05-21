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
from ...utils.console import console
from ...utils.errors import handle_graph_errors


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

    query_params = TrendingRequestBuilder.TrendingRequestBuilderGetQueryParameters(
        top=min(args.top, 100),
    )
    config = RequestConfiguration(query_parameters=query_params)

    result = await context.graph_client.me.insights.trending.get(
        request_configuration=config
    )

    items = result.value if result and result.value else []

    if not items:
        logger.info("No trending items found.")
        return 0

    # Client-side type filter
    type_filter = args.type.lower() if args.type else None
    if type_filter:
        items = [
            i
            for i in items
            if i.resource_visualization
            and i.resource_visualization.type
            and type_filter in i.resource_visualization.type.lower()
        ]

    table = Table(show_header=True, header_style="bold", box=None, padding=(0, 1))
    table.add_column("#", style="dim", justify="right", width=4)
    table.add_column("Type", style="cyan", width=12)
    table.add_column("Title", min_width=40)
    table.add_column("Site", style="dim", max_width=30)
    table.add_column("Weight", style="yellow", justify="right", width=8)

    for i, item in enumerate(items, 1):
        viz = item.resource_visualization
        ref = item.resource_reference

        title = (viz.title if viz else None) or "(untitled)"
        file_type = (viz.type if viz else None) or "?"
        site = (viz.container_display_name if viz else None) or ""
        weight = f"{item.weight:.3f}" if item.weight else ""

        # Truncate long titles
        if len(title) > 80:
            title = title[:77] + "..."

        table.add_row(str(i), file_type, title, site, weight)

    console.print("[bold]Trending around you[/bold]")
    console.rule()
    console.print(table)
    logger.success(f"{len(items)} trending item(s).")

    return 0
