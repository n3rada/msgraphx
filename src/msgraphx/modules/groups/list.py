# msgraphx/modules/groups/list.py
#
# Enumerate M365 Unified groups.
#
# --mine (delegated):  GET /me/transitiveMemberOf/microsoft.graph.group
#   Covers direct AND nested group membership. The only reliable way to know
#   every group that grants the user access to SharePoint sites, Teams, and
#   shared mailboxes — a simple /me/memberOf misses nested/transitive membership.
#
# default (any token): GET /groups?$filter=groupTypes/any(c:c eq 'Unified')
#   App-only: tenant-wide. Delegated: all groups the token can read.

from __future__ import annotations

# Built-in imports
import argparse
from typing import Any

# External library imports
from kiota_abstractions.base_request_configuration import RequestConfiguration
from loguru import logger
from msgraph.generated.groups.groups_request_builder import GroupsRequestBuilder
from msgraph.generated.models.group import Group
from rich.table import Table

# Local library imports
from ...core.context import GraphContext
from ...utils import cache, output
from ...utils.console import console
from ...utils.errors import handle_graph_errors
from ...utils.pagination import GraphPaginator, collect_all
from ..sharepoint.groups import get_user_m365_groups
from ...utils.roles import require_scopes

def _to_dict(g: Group) -> dict[str, Any]:
    """Convert a group into the JSON-serializable row used for output and cache.

    Args:
        g: The group to convert.

    Returns:
        A dict with `id`, `display_name`, `mail`, `visibility`,
        `teams_provisioned`, `description`, and `created_datetime` (ISO 8601
        string or None). Missing string fields become "".
    """
    return {
        "id": g.id,
        "display_name": g.display_name or "",
        "mail": g.mail or "",
        "visibility": g.visibility or "",
        "teams_provisioned": bool(
            g.resource_provisioning_options
            and "Team" in g.resource_provisioning_options
        ),
        "description": getattr(g, "description", None) or "",
        "created_datetime": (
            g.created_date_time.isoformat()
            if g.created_date_time
            else None
        ),
    }

async def fetch_mine(
    context: GraphContext,
    visibility: str | None = None,
    teams_only: bool = False,
) -> list[dict[str, Any]]:
    """Return the signed-in user's M365 Unified groups, including nested ones.

    Requires a delegated token.

    Args:
        context: Shared Graph runtime context.
        visibility: "Private" or "Public" to keep only that visibility. None
            returns both.
        teams_only: When True, keep only groups with a Teams workspace.

    Returns:
        One row per group, as built by `_to_dict`.
    """
    groups = await get_user_m365_groups(
        context.graph_client,
        visibility=visibility,
        teams_only=teams_only,
    )
    return [_to_dict(g) for g in groups]

async def fetch_all(
    context: GraphContext,
    visibility: str | None = None,
    teams_only: bool = False,
) -> list[dict[str, Any]]:
    """Return every M365 Unified group the token can read.

    App-only tokens see the whole tenant; delegated tokens see the groups the
    user is allowed to read.

    Args:
        context: Shared Graph runtime context.
        visibility: "Private" or "Public" to keep only that visibility. None
            returns both.
        teams_only: When True, keep only groups with a Teams workspace.

    Returns:
        One row per group, as built by `_to_dict`.
    """
    query_params = GroupsRequestBuilder.GroupsRequestBuilderGetQueryParameters(
        filter="groupTypes/any(c:c eq 'Unified')",
        select=[
            "id", "displayName", "mail", "visibility", "description",
            "resourceProvisioningOptions", "createdDateTime", "groupTypes",
        ],
        top=999,
    )
    config = RequestConfiguration(query_parameters=query_params)
    all_groups: list[Group] = await collect_all(context.graph_client.groups, config)

    if teams_only:
        all_groups = [
            g for g in all_groups
            if g.resource_provisioning_options
            and "Team" in g.resource_provisioning_options
        ]
    if visibility:
        all_groups = [g for g in all_groups if g.visibility == visibility]

    return [_to_dict(g) for g in all_groups]

def add_arguments(parser: argparse.ArgumentParser) -> None:
    """Register the `groups list` arguments on a parser.

    Args:
        parser: Parser for the `list` subcommand.
    """
    parser.add_argument(
        "--mine",
        action="store_true",
        help="Only show groups the current user belongs to (direct and transitive). Requires delegated auth.",
    )
    parser.add_argument(
        "--teams-only",
        action="store_true",
        help="Only show groups that have a Microsoft Teams workspace provisioned.",
    )
    parser.add_argument(
        "--visibility",
        choices=["Private", "Public"],
        default=None,
        help="Filter by visibility.",
    )

@handle_graph_errors
@require_scopes("Group.Read.All")
async def run_with_arguments(context: GraphContext, args: argparse.Namespace) -> int:
    """List M365 Unified groups and cache them.

    With `--mine`, lists the user's transitive memberships; otherwise lists
    every readable Unified group. Results are cached under the `groups` key
    and printed as a table, JSON, or NDJSON depending on the output mode.

    Args:
        context: Shared Graph runtime context.
        args: Parsed arguments (`mine`, `teams_only`, `visibility`).

    Returns:
        0 on success (including no groups found), 1 when `--mine` is used
        with an app-only token.

    Raises:
        AuthenticationError: The Graph token is invalid or expired.
    """
    mine = getattr(args, "mine", False)
    teams_only = getattr(args, "teams_only", False)
    visibility = getattr(args, "visibility", None)

    if mine:
        if context.is_app_only:
            logger.error("--mine requires delegated authentication (user context).")
            return 1
        logger.info("Fetching current user's M365 Unified groups (transitive)")
        rows = await fetch_mine(context, visibility=visibility, teams_only=teams_only)
    else:
        logger.info("Enumerating all M365 Unified groups")
        rows = await fetch_all(context, visibility=visibility, teams_only=teams_only)

    if not rows:
        logger.info("No M365 Unified groups found.")
        if context.json_output:
            output.print_json([])
        return 0

    cache.save_results(rows, key="groups", identity=context.identity_hash)

    if context.json_output:
        output.print_json(rows)
        return 0

    if context.ndjson_output:
        for row in rows:
            output.print_ndjson_item(row)
        return 0

    private_count = sum(1 for r in rows if r["visibility"] == "Private")
    title = (
        f"My M365 Unified Groups ({len(rows)}, {private_count} private)"
        if mine
        else f"M365 Unified Groups ({len(rows)}, {private_count} private)"
    )

    table = Table(title=title, show_header=True, header_style="bold", box=None, padding=(0, 1))
    table.add_column("#", justify="right", style="dim")
    table.add_column("Group")
    table.add_column("Email", style="dim")
    table.add_column("Visibility")
    table.add_column("Teams", justify="center")

    for idx, row in enumerate(rows, 1):
        vis = row["visibility"]
        vis_display = f"[red]{vis}[/red]" if vis == "Private" else f"[green]{vis}[/green]"
        table.add_row(
            str(idx),
            row["display_name"] or "(unnamed)",
            row["mail"],
            vis_display,
            "yes" if row["teams_provisioned"] else "",
        )

    console.print(table)
    logger.success(f"{len(rows)} group(s) found")
    return 0
