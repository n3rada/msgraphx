# msgraphx/modules/me/planner.py
#
# List Planner tasks assigned to the current user.
# Task titles and plan names reveal active projects and internal priorities.
#
# Required delegated permissions:
#   Tasks.Read

# Built-in imports
from __future__ import annotations

import argparse

# External library imports
from loguru import logger
from rich.table import Table

# Local library imports
from ...core.context import GraphContext
from ...utils import output
from ...utils.console import console
from ...utils.errors import handle_graph_errors
from ...utils.pagination import collect_all


def add_arguments(parser: "argparse.ArgumentParser") -> None:
    parser.add_argument(
        "--top",
        "-n",
        type=int,
        default=50,
        metavar="N",
        help="Maximum number of tasks to return (default: 50).",
    )


async def fetch(context: GraphContext, top: int = 50) -> list[dict]:
    """Return Planner tasks for the current user as plain dicts.

    Raises on API error — callers are responsible for handling exceptions.
    """
    tasks = await collect_all(context.graph_client.me.planner.tasks)
    tasks = tasks[:top]

    rows = []
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
async def run_with_arguments(
    context: "GraphContext", args: "argparse.Namespace"
) -> int:
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
