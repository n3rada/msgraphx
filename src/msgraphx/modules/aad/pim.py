# msgraphx/modules/aad/pim.py
#
# Enumerate active and eligible PIM role assignments via the Graph SDK.
#
# Required permissions:
#   RoleManagement.Read.Directory (or RoleManagement.ReadWrite.Directory)

from __future__ import annotations

# Built-in imports
import argparse
import asyncio

# External library imports
from kiota_abstractions.base_request_configuration import RequestConfiguration
from loguru import logger
from msgraph.generated.role_management.directory.role_assignment_schedule_instances.role_assignment_schedule_instances_request_builder import (
    RoleAssignmentScheduleInstancesRequestBuilder,
)
from msgraph.generated.role_management.directory.role_eligibility_schedule_instances.role_eligibility_schedule_instances_request_builder import (
    RoleEligibilityScheduleInstancesRequestBuilder,
)
from rich.table import Table

# Local library imports
from ...core.context import GraphContext
from ...utils import output, pagination
from ...utils.console import console
from ...utils.errors import handle_graph_errors, raise_if_forbidden


def _principal_label(item) -> str:
    principal = getattr(item, "principal", None)
    if principal:
        ad = getattr(principal, "additional_data", {}) or {}
        upn = ad.get("userPrincipalName")
        if upn:
            return upn
        name = ad.get("displayName") or getattr(principal, "display_name", None)
        if name:
            return name
    return getattr(item, "principal_id", None) or ""


def _role_label(item) -> str:
    rd = getattr(item, "role_definition", None)
    if rd:
        return getattr(rd, "display_name", None) or getattr(item, "role_definition_id", None) or ""
    return getattr(item, "role_definition_id", None) or ""


def _fmt_dt(dt) -> str:
    if not dt:
        return ""
    return dt.strftime("%Y-%m-%d")


async def _safe_collect(builder, config) -> list:
    try:
        return await pagination.collect_all(builder, request_configuration=config)
    except Exception as exc:
        raise_if_forbidden(exc)
        logger.warning(f"Could not fetch PIM assignments: {exc}")
        return []


async def fetch(context: GraphContext, state: str) -> list[dict]:
    """Return active and/or eligible PIM role assignments as plain dicts."""
    ActiveParams = RoleAssignmentScheduleInstancesRequestBuilder.RoleAssignmentScheduleInstancesRequestBuilderGetQueryParameters
    EligibleParams = RoleEligibilityScheduleInstancesRequestBuilder.RoleEligibilityScheduleInstancesRequestBuilderGetQueryParameters

    active_config = RequestConfiguration(
        query_parameters=ActiveParams(expand=["principal", "roleDefinition"])
    )
    eligible_config = RequestConfiguration(
        query_parameters=EligibleParams(expand=["principal", "roleDefinition"])
    )

    active_builder = context.graph_client.role_management.directory.role_assignment_schedule_instances
    eligible_builder = context.graph_client.role_management.directory.role_eligibility_schedule_instances

    if state == "active":
        active_items = await _safe_collect(active_builder, active_config)
        eligible_items = []
    elif state == "eligible":
        active_items = []
        eligible_items = await _safe_collect(eligible_builder, eligible_config)
    else:
        active_items, eligible_items = await asyncio.gather(
            _safe_collect(active_builder, active_config),
            _safe_collect(eligible_builder, eligible_config),
        )

    rows = []
    for item in active_items:
        rows.append({
            "assignment_state": "Active",
            "assignment_type": getattr(item, "assignment_type", None),
            "role_display_name": _role_label(item),
            "role_definition_id": getattr(item, "role_definition_id", None),
            "subject_id": getattr(item, "principal_id", None),
            "subject": _principal_label(item),
            "member_type": getattr(item, "member_type", None),
            "start_date": _fmt_dt(getattr(item, "start_date_time", None)),
            "end_date": _fmt_dt(getattr(item, "end_date_time", None)),
        })
    for item in eligible_items:
        rows.append({
            "assignment_state": "Eligible",
            "assignment_type": None,
            "role_display_name": _role_label(item),
            "role_definition_id": getattr(item, "role_definition_id", None),
            "subject_id": getattr(item, "principal_id", None),
            "subject": _principal_label(item),
            "member_type": getattr(item, "member_type", None),
            "start_date": _fmt_dt(getattr(item, "start_date_time", None)),
            "end_date": _fmt_dt(getattr(item, "end_date_time", None)),
        })
    return rows


def add_arguments(parser: argparse.ArgumentParser) -> None:
    parser.add_argument(
        "--state",
        choices=["active", "eligible", "all"],
        default="all",
        help="Assignment state to show (default: all).",
    )


@handle_graph_errors
async def run_with_arguments(context: GraphContext, args: argparse.Namespace) -> int:
    logger.info("Fetching PIM role assignments")
    rows = await fetch(context, args.state)

    if not rows:
        logger.info("No PIM role assignments found.")
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
    table.add_column("State", width=10)
    table.add_column("Role", min_width=35)
    table.add_column("Subject", min_width=30)
    table.add_column("Type", style="dim", width=12)
    table.add_column("Start", style="dim", width=11)
    table.add_column("End", style="dim", width=11)

    state_colors = {
        "Active": "[green]Active[/green]",
        "Eligible": "[yellow]Eligible[/yellow]",
    }

    for row in rows:
        state_display = state_colors.get(row["assignment_state"], row["assignment_state"])
        table.add_row(
            state_display,
            row["role_display_name"] or "",
            row["subject"] or "",
            row["member_type"] or "",
            row["start_date"],
            row["end_date"],
        )

    console.print("[bold]PIM role assignments[/bold]")
    console.rule()
    console.print(table)
    logger.success(f"{len(rows)} assignment(s) found.")
    return 0
