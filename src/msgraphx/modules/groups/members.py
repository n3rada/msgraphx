# msgraphx/modules/groups/members.py
#
# List members of an M365 Unified group.
#
# Default: GET /groups/{id}/members — direct members only.
# --transitive: GET /groups/{id}/transitiveMembers — all members including
#   those inherited through nested groups. Essential for understanding who
#   actually has access when groups are nested.
#
# Members can be Users, Groups, Devices, or ServicePrincipals.
# Each type is normalised into a flat dict with a `type` discriminator.

from __future__ import annotations

import argparse

from loguru import logger
from msgraph.generated.models.user import User
from msgraph.generated.models.group import Group
from msgraph.generated.models.device import Device
from msgraph.generated.models.service_principal import ServicePrincipal
from rich.table import Table

from ...core.context import GraphContext
from ...utils import cache, output
from ...utils.console import console
from ...utils.errors import handle_graph_errors
from ...utils.pagination import GraphPaginator
from ...utils.roles import require_scopes

def _member_to_dict(m) -> dict:
    if isinstance(m, User):
        return {
            "type": "user",
            "id": m.id,
            "display_name": m.display_name or "",
            "identifier": m.user_principal_name or m.mail or "",
            "account_enabled": m.account_enabled,
            "job_title": m.job_title or "",
            "department": m.department or "",
        }
    if isinstance(m, Group):
        return {
            "type": "group",
            "id": m.id,
            "display_name": m.display_name or "",
            "identifier": m.mail or "",
            "account_enabled": None,
            "job_title": "",
            "department": "",
        }
    if isinstance(m, ServicePrincipal):
        return {
            "type": "service_principal",
            "id": m.id,
            "display_name": m.display_name or "",
            "identifier": m.app_id or "",
            "account_enabled": m.account_enabled,
            "job_title": "",
            "department": "",
        }
    if isinstance(m, Device):
        return {
            "type": "device",
            "id": m.id,
            "display_name": m.display_name or "",
            "identifier": m.device_id or "",
            "account_enabled": m.account_enabled,
            "job_title": "",
            "department": "",
        }
    return {
        "type": "unknown",
        "id": m.id or "",
        "display_name": getattr(m, "display_name", "") or "",
        "identifier": "",
        "account_enabled": None,
        "job_title": "",
        "department": "",
    }

def add_arguments(parser: argparse.ArgumentParser) -> None:
    parser.add_argument(
        "group_id",
        help="Group ID (GUID) or mail address to look up members for.",
    )
    parser.add_argument(
        "--transitive",
        action="store_true",
        help="Include members inherited through nested groups (full transitive membership).",
    )
    parser.add_argument(
        "--users-only",
        action="store_true",
        help="Only show user members, skip groups and service principals.",
    )

@handle_graph_errors
@require_scopes("GroupMember.Read.All")
async def run_with_arguments(context: GraphContext, args: argparse.Namespace) -> int:
    group_id = args.group_id
    transitive = getattr(args, "transitive", False)
    users_only = getattr(args, "users_only", False)

    endpoint = "transitiveMembers" if transitive else "members"
    logger.info(f"Fetching {endpoint} for group {group_id}")

    builder = (
        context.graph_client.groups.by_group_id(group_id).transitive_members
        if transitive
        else context.graph_client.groups.by_group_id(group_id).members
    )

    rows = []
    async for member in GraphPaginator(builder):
        rows.append(_member_to_dict(member))

    if users_only:
        rows = [r for r in rows if r["type"] == "user"]

    if not rows:
        logger.info("No members found.")
        if context.json_output:
            output.print_json([])
        return 0

    cache.save_results(rows, key="group_members", identity=context.identity_hash)

    if context.json_output:
        output.print_json(rows)
        return 0

    if context.ndjson_output:
        for row in rows:
            output.print_ndjson_item(row)
        return 0

    type_counts = {}
    for r in rows:
        type_counts[r["type"]] = type_counts.get(r["type"], 0) + 1
    summary = ", ".join(f"{v} {k}(s)" for k, v in sorted(type_counts.items()))

    label = "transitive members" if transitive else "direct members"
    table = Table(
        title=f"Group {label} ({len(rows)} — {summary})",
        show_header=True,
        header_style="bold",
        box=None,
        padding=(0, 1),
    )
    table.add_column("#", justify="right", style="dim")
    table.add_column("Type", style="dim", width=16)
    table.add_column("Display name")
    table.add_column("Identifier", style="dim")

    type_styles = {
        "user": "cyan",
        "group": "yellow",
        "service_principal": "magenta",
        "device": "blue",
    }

    for idx, row in enumerate(rows, 1):
        t = row["type"]
        style = type_styles.get(t, "")
        table.add_row(
            str(idx),
            f"[{style}]{t}[/{style}]" if style else t,
            row["display_name"],
            row["identifier"],
        )

    console.print(table)
    logger.success(f"{len(rows)} member(s) found")
    return 0
