# msgraphx/modules/me/used.py
#
# List documents recently used by the current user via the Insights API.
# Shows files the user recently opened from OneDrive, SharePoint, or Teams.
#
# NOTE: This API is deprecated and will stop returning data after 2028-01-01.

# Built-in imports
from __future__ import annotations

import argparse
import warnings
from typing import Any

# External library imports
from kiota_abstractions.base_request_configuration import RequestConfiguration
from loguru import logger
from msgraph.generated.users.item.insights.used.used_request_builder import (
    UsedRequestBuilder,
)
from rich.table import Table

# Local library imports
from ...core.context import GraphContext
from ...utils import output
from ...utils.console import console
from ...utils.errors import handle_graph_errors
from ...utils.roles import require_scopes

def add_arguments(parser: argparse.ArgumentParser) -> None:
    """Register the `me used` options on the parser.

    Args:
        parser: Parser for the `me used` subcommand.
    """
    parser.add_argument(
        "--top",
        "-n",
        type=int,
        default=25,
        metavar="N",
        help="Number of items to display (default: 25).",
    )

async def fetch(context: GraphContext, top: int = 25) -> list[dict[str, Any]]:
    """Return documents the current user recently used as plain dicts.

    Uses the deprecated `/me/insights/used` API; the SDK's deprecation
    warning is suppressed. Only the first page of results is read.

    Args:
        context: Shared Graph runtime context (delegated auth).
        top: Maximum number of items to request (capped at 100).

    Returns:
        One dict per item with the keys `type`, `title`, `site`,
        `last_used` (`YYYY-MM-DD`), `url`, and `id`.

    Raises:
        ODataError: If the Graph API rejects the request.
    """
    query_params = UsedRequestBuilder.UsedRequestBuilderGetQueryParameters(
        top=min(top, 100),
    )
    config = RequestConfiguration(query_parameters=query_params)

    with warnings.catch_warnings():
        warnings.simplefilter("ignore", DeprecationWarning)
        result = await context.graph_client.me.insights.used.get(
            request_configuration=config
        )

    items = (result.value or []) if result else []

    rows: list[dict[str, Any]] = []
    for item in items:
        viz = item.resource_visualization
        ref = item.resource_reference
        last_used = ""
        if item.last_used and item.last_used.last_accessed_date_time:
            last_used = item.last_used.last_accessed_date_time.strftime("%Y-%m-%d")

        rows.append({
            "type": (viz.type if viz else None),
            "title": (viz.title if viz else None),
            "site": (viz.container_display_name if viz else None),
            "last_used": last_used,
            "url": (ref.web_url if ref else None),
            "id": (ref.id if ref else None),
        })

    return rows

@handle_graph_errors
@require_scopes("Sites.Read.All")
async def run_with_arguments(
    context: GraphContext, args: argparse.Namespace
) -> int:
    """List documents the current user recently used.

    Prints a table, or JSON/NDJSON when the matching output flag is set.

    Args:
        context: Shared Graph runtime context.
        args: Parsed CLI arguments (`top`).

    Returns:
        0 on success (including no items), 1 for app-only auth.

    Raises:
        AuthenticationError: If the Graph token is invalid or expired.
        ODataError: If the Graph API rejects the request.
    """
    if context.is_app_only:
        logger.error("This module requires delegated authentication (user context).")
        return 1

    logger.info("Fetching recently used documents")

    rows = await fetch(context, top=args.top)

    if not rows:
        logger.info("No recently used items found.")
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
    table.add_column("Site", style="dim", min_width=25)
    table.add_column("Last used", style="dim", width=12)

    for i, row in enumerate(rows, 1):
        table.add_row(
            str(i),
            row["type"] or "?",
            row["title"] or "(untitled)",
            row["site"] or "",
            row["last_used"],
        )

    console.print("[bold]Recently used documents[/bold]")
    console.rule()
    console.print(table)
    logger.success(f"{len(rows)} item(s) found.")
    return 0
