# msgraphx/modules/aad/ca.py
#
# List conditional access policies in the tenant.
# Shows which apps, users, and conditions each policy targets
# and what grant controls it requires.
#
# Required delegated permissions:
#   Policy.Read.All  (admin consent required)

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


async def fetch(context: GraphContext, state: str | None = None) -> list[dict]:
    """Return conditional access policies as plain dicts.

    `state` accepts 'enabled', 'disabled', or 'report'.
    Raises on API error. Callers are responsible for handling exceptions.
    """
    policies = await collect_all(context.graph_client.policies.conditional_access_policies)

    state_map = {
        "enabled": "enabled",
        "disabled": "disabled",
        "report": "enabledForReportingButNotEnforced",
    }
    if state:
        target_state = state_map.get(state, state)
        policies = [
            p for p in policies
            if str(p.state).split(".")[-1].lower() == target_state.lower()
        ]

    rows = []
    for p in policies:
        state_str = str(p.state).split(".")[-1] if p.state else "unknown"

        included_users = []
        excluded_users = []
        included_apps = []
        if p.conditions:
            if p.conditions.users:
                included_users = p.conditions.users.include_users or []
                excluded_users = p.conditions.users.exclude_users or []
                inc_groups = p.conditions.users.include_groups or []
                if inc_groups:
                    included_users = included_users + [f"groups:{len(inc_groups)}"]
            if p.conditions.applications:
                included_apps = p.conditions.applications.include_applications or []

        grant_controls: list[str] = []
        if p.grant_controls and p.grant_controls.built_in_controls:
            grant_controls = [
                str(c).split(".")[-1] for c in p.grant_controls.built_in_controls
            ]

        rows.append({
            "id": p.id,
            "display_name": p.display_name,
            "state": state_str,
            "include_users": included_users,
            "exclude_users": excluded_users,
            "include_apps": included_apps,
            "grant_controls": grant_controls,
            "created": (
                p.created_date_time.strftime("%Y-%m-%d") if p.created_date_time else ""
            ),
            "modified": (
                p.modified_date_time.strftime("%Y-%m-%d") if p.modified_date_time else ""
            ),
        })

    return rows


def add_arguments(parser: "argparse.ArgumentParser") -> None:
    parser.add_argument(
        "--state",
        choices=["enabled", "disabled", "report"],
        default=None,
        help="Filter by policy state: enabled, disabled, or report (reporting only).",
    )


@handle_graph_errors
async def run_with_arguments(
    context: "GraphContext", args: "argparse.Namespace"
) -> int:
    logger.info("Fetching conditional access policies")

    rows = await fetch(context, state=args.state)

    if not rows:
        logger.info("No conditional access policies found.")
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
    table.add_column("Name", min_width=35)
    table.add_column("State", style="cyan", width=10)
    table.add_column("Grant controls", style="dim", min_width=20)
    table.add_column("Modified", style="dim", width=12)

    state_colors = {
        "enabled": "[green]enabled[/green]",
        "disabled": "[red]disabled[/red]",
        "enabledForReportingButNotEnforced": "[yellow]report[/yellow]",
    }

    for i, row in enumerate(rows, 1):
        state_display = state_colors.get(row["state"], row["state"])
        controls = ", ".join(row["grant_controls"]) if row["grant_controls"] else "block"
        table.add_row(str(i), row["display_name"] or "", state_display, controls, row["modified"])

    console.print("[bold]Conditional access policies[/bold]")
    console.rule()
    console.print(table)
    logger.success(f"{len(rows)} policy/policies found.")
    return 0
