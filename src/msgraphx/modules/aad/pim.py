# msgraphx/modules/aad/pim.py
#
# Enumerate active and eligible PIM role assignments.
# Uses api.azrbac.mspim.azure.com, which is the legacy Azure RBAC PIM API
# (not the Graph API). The tenant resource ID is discovered automatically.
#
# Required delegated permissions:
#   PrivilegedAccess.Read.AzureAD  (or PrivilegedAccess.ReadWrite.AzureAD)

from __future__ import annotations

import argparse
from urllib.parse import quote_plus

# External library imports
import httpx
from loguru import logger
from rich.table import Table

# Local library imports
from ...core.context import GraphContext
from ...utils import output
from ...utils.console import console
from ...utils.errors import handle_graph_errors

_BASE = "https://api.azrbac.mspim.azure.com/api/v2/privilegedAccess/aadroles"


async def _get_resource_id(token: str) -> str | None:
    async with httpx.AsyncClient(verify=False, timeout=30) as client:
        resp = await client.get(
            f"{_BASE}/resources?$select=id,displayName,type,externalId&$expand=parent",
            headers={"Authorization": f"Bearer {token}"},
        )
    if resp.status_code != 200:
        logger.error(f"Failed to get PIM resource ID: {resp.status_code}")
        return None
    values = resp.json().get("value", [])
    return values[0]["id"] if values else None


async def _fetch_by_state(token: str, resource_id: str, state: str) -> list[dict]:
    rid = quote_plus(resource_id)
    uri = (
        f"{_BASE}/roleAssignments"
        f"?$expand=linkedEligibleRoleAssignment,subject,scopedResource,"
        f"roleDefinition($expand=resource)"
        f"&$count=true"
        f"&$filter=(roleDefinition/resource/id%20eq%20%27{rid}%27)"
        f"%20and%20(assignmentState%20eq%20%27{state}%27)"
        f"&$orderby=roleDefinition/displayName"
        f"&$skip=0&$top=10000"
    )
    results = []
    async with httpx.AsyncClient(verify=False, timeout=30) as client:
        while uri:
            resp = await client.get(uri, headers={"Authorization": f"Bearer {token}"})
            if resp.status_code != 200:
                logger.error(f"PIM {state} fetch failed: {resp.status_code}")
                break
            data = resp.json()
            results.extend(data.get("value", []))
            uri = data.get("@odata.nextLink")
    return results


async def fetch(context: GraphContext) -> list[dict]:
    """Return all active and eligible PIM role assignments as plain dicts."""
    token = await context.get_access_token()
    if not token:
        raise RuntimeError("No access token available.")

    resource_id = await _get_resource_id(token)
    if not resource_id:
        raise RuntimeError("Could not determine PIM resource ID. Check PrivilegedAccess.Read.AzureAD scope.")

    logger.debug(f"PIM resource ID: {resource_id}")
    active = await _fetch_by_state(token, resource_id, "Active")
    eligible = await _fetch_by_state(token, resource_id, "Eligible")

    rows = []
    for a in active + eligible:
        role_def = a.get("roleDefinition") or {}
        subject = a.get("subject") or {}
        rows.append({
            "id": a.get("id"),
            "assignment_state": a.get("assignmentState"),
            "role_display_name": role_def.get("displayName"),
            "role_id": role_def.get("id"),
            "subject_id": subject.get("id"),
            "subject_display_name": subject.get("displayName"),
            "subject_upn": subject.get("userPrincipalName"),
            "subject_type": subject.get("type"),
            "is_permanent": a.get("isPermanent"),
            "start_date": a.get("startDateTime"),
            "end_date": a.get("endDateTime"),
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
    rows = await fetch(context)

    if args.state != "all":
        rows = [r for r in rows if (r["assignment_state"] or "").lower() == args.state]

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
    table.add_column("Type", style="dim", width=10)
    table.add_column("Permanent", style="dim", width=9)

    state_colors = {
        "Active": "[green]Active[/green]",
        "Eligible": "[yellow]Eligible[/yellow]",
    }

    for row in rows:
        state_display = state_colors.get(row["assignment_state"] or "", row["assignment_state"] or "")
        subject = row["subject_upn"] or row["subject_display_name"] or row["subject_id"] or ""
        perm = "yes" if row["is_permanent"] else "no"
        table.add_row(
            state_display,
            row["role_display_name"] or "",
            subject,
            row["subject_type"] or "",
            perm,
        )

    console.print("[bold]PIM role assignments[/bold]")
    console.rule()
    console.print(table)
    logger.success(f"{len(rows)} assignment(s) found.")
    return 0
