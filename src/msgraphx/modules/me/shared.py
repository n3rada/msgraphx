# msgraphx/modules/me/shared.py
#
# List documents shared with the current user via the Insights API.
# Complements 'me trending': shows files others explicitly shared with you
# rather than items Microsoft infers are relevant.
#
# NOTE: This API is deprecated and will stop returning data after 2028-01-01.
#
# Required delegated permissions:
#   Sites.Read.All

# Built-in imports
from __future__ import annotations

import argparse
import warnings

# External library imports
from kiota_abstractions.base_request_configuration import RequestConfiguration
from loguru import logger
from msgraph.generated.users.item.insights.shared.shared_request_builder import (
    SharedRequestBuilder,
)
from rich.table import Table

# Local library imports
from ...core.context import GraphContext
from ...utils import output
from ...utils.console import console
from ...utils.errors import handle_graph_errors
from ...utils.roles import require_scopes


def add_arguments(parser: "argparse.ArgumentParser") -> None:
    parser.add_argument(
        "--top",
        "-n",
        type=int,
        default=25,
        metavar="N",
        help="Number of items to display (default: 25).",
    )


async def fetch(context: GraphContext, top: int = 25) -> list[dict]:
    """Return documents shared with the current user as plain dicts.

    Raises on API error. Callers are responsible for handling exceptions.
    """
    query_params = SharedRequestBuilder.SharedRequestBuilderGetQueryParameters(
        top=min(top, 100),
    )
    config = RequestConfiguration(query_parameters=query_params)

    with warnings.catch_warnings():
        warnings.simplefilter("ignore", DeprecationWarning)
        result = await context.graph_client.me.insights.shared.get(
            request_configuration=config
        )

    items = (result.value or []) if result else []

    rows = []
    for item in items:
        viz = item.resource_visualization
        ref = item.resource_reference
        shared_by = ""
        shared_date = ""
        if item.last_shared:
            if item.last_shared.shared_by:
                shared_by = item.last_shared.shared_by.display_name or ""
            if item.last_shared.shared_date_time:
                shared_date = item.last_shared.shared_date_time.strftime("%Y-%m-%d")

        rows.append({
            "type": (viz.type if viz else None),
            "title": (viz.title if viz else None),
            "site": (viz.container_display_name if viz else None),
            "shared_by": shared_by,
            "shared_date": shared_date,
            "url": (ref.web_url if ref else None),
            "id": (ref.id if ref else None),
        })

    return rows


@handle_graph_errors
@require_scopes("Sites.Read.All")
async def run_with_arguments(
    context: "GraphContext", args: "argparse.Namespace"
) -> int:
    if context.is_app_only:
        logger.error("This module requires delegated authentication (user context).")
        return 1

    logger.info("Fetching shared documents")

    rows = await fetch(context, top=args.top)

    if not rows:
        logger.info("No shared items found.")
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
    table.add_column("Shared by", style="dim", min_width=20)
    table.add_column("Date", style="dim", width=12)

    for i, row in enumerate(rows, 1):
        table.add_row(
            str(i),
            row["type"] or "?",
            row["title"] or "(untitled)",
            row["shared_by"],
            row["shared_date"],
        )

    console.print("[bold]Shared with you[/bold]")
    console.rule()
    console.print(table)
    logger.success(f"{len(rows)} item(s) found.")
    return 0
