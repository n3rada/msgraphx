# msgraphx/modules/me/planner.py
#
# List Planner tasks assigned to the current user.
# Task titles and plan names reveal active projects and internal priorities.

# Built-in imports
from __future__ import annotations

import argparse
from typing import Any

# External library imports
from loguru import logger
from msgraph.generated.models.planner_task import PlannerTask
from rich.table import Table

# Local library imports
from ...core.context import GraphContext
from ...utils import output
from ...utils.console import console
from ...utils.errors import handle_graph_errors
from ...utils.pagination import collect_all
from ...utils.roles import require_scopes

def add_arguments(parser: argparse.ArgumentParser) -> None:
    """Register the `me planner` options on the parser.

    Args:
        parser: Parser for the `me planner` subcommand.
    """
    parser.add_argument(
        "--top",
        "-n",
        type=int,
        default=50,
        metavar="N",
        help="Maximum number of tasks to return (default: 50).",
    )

async def fetch(context: GraphContext, top: int = 50) -> list[dict[str, Any]]:
    """Return the Planner tasks assigned to the current user as plain dicts.

    All pages are fetched, then truncated to `top`.

    Args:
        context: Shared Graph runtime context (delegated auth).
        top: Maximum number of tasks to return.

    Returns:
        One dict per task with the keys `id`, `title`, `plan_id`,
        `bucket_id`, `percent_complete` (0 when unset), `due`, and
        `created` (dates as `YYYY-MM-DD`, empty when unset).

    Raises:
        AuthenticationError: If the Graph token is invalid or expired.
        ODataError: If the Graph API rejects the request.
    """
    tasks: list[PlannerTask] = await collect_all(context.graph_client.me.planner.tasks)
    tasks = tasks[:top]

    rows: list[dict[str, Any]] = []
    for task in tasks:
        due = ""
        if task.due_date_time:
            due = task.due_date_time.strftime("%Y-%m-%d")

        created = ""
        if task.created_date_time:
            created = task.created_date_time.strftime("%Y-%m-%d")

        rows.append({
            "id": task.id,
            "title": task.title,
            "plan_id": task.plan_id,
            "bucket_id": task.bucket_id,
            "percent_complete": task.percent_complete or 0,
            "due": due,
            "created": created,
        })

    return rows

@handle_graph_errors
@require_scopes("Tasks.Read")
async def run_with_arguments(
    context: GraphContext, args: argparse.Namespace
) -> int:
    """List the Planner tasks assigned to the current user.

    Prints a table, or JSON/NDJSON when the matching output flag is set.

    Args:
        context: Shared Graph runtime context.
        args: Parsed CLI arguments (`top`).

    Returns:
        0 on success (including no tasks), 1 for app-only auth.

    Raises:
        AuthenticationError: If the Graph token is invalid or expired.
        ODataError: If the Graph API rejects the request.
    """
    if context.is_app_only:
        logger.error("This module requires delegated authentication (user context).")
        return 1

    logger.info("Fetching Planner tasks")

    rows = await fetch(context, top=args.top)

    if not rows:
        logger.info("No Planner tasks found.")
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
    table.add_column("Task", min_width=40)
    table.add_column("Done%", justify="right", width=6)
    table.add_column("Due", style="cyan", width=12)
    table.add_column("Plan ID", style="dim", width=38)

    for i, row in enumerate(rows, 1):
        table.add_row(
            str(i),
            row["title"] or "",
            str(row["percent_complete"]),
            row["due"],
            row["plan_id"] or "",
        )

    console.print("[bold]Planner tasks[/bold]")
    console.rule()
    console.print(table)
    logger.success(f"{len(rows)} task(s) found.")
    return 0
