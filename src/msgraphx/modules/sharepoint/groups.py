# msgraphx/modules/sharepoint/groups.py
#
# Helper and CLI subcommand for fetching the current user's M365 Unified groups.
#
# Background: Microsoft 365 group types
# --------------------------------------
# Azure AD has several group types. The relevant distinction here:
#
#   Security group     — used for access control only; no shared resources.
#   Distribution group — email distribution list; no shared resources.
#   M365 Unified group — the collaboration backbone. Every Unified group
#                        automatically provisions a SharePoint team site,
#                        a shared mailbox, a shared calendar, and a Planner board.
#
# A Microsoft Teams workspace is an *optional* layer that can be added on top of
# a Unified group. Its presence is signalled by "Team" appearing in the group's
# `resourceProvisioningOptions` field. Not every Unified group has Teams:
# groups created via SharePoint admin, Planner, or Exchange may never have
# Teams provisioned. Those groups still own a SharePoint site and are fully
# relevant for file enumeration — filtering to teams_only would silently drop them.
#
# `GET /me/transitiveMemberOf/microsoft.graph.group` returns all groups
# (Security, Distribution, Unified) the user belongs to, including via nested
# membership. We filter to `groupTypes == ["Unified"]` to keep only M365 groups.
#
# Required delegated permissions: Group.Read.All (or GroupMember.Read.All)

from __future__ import annotations

import argparse

from loguru import logger
from msgraph.graph_service_client import GraphServiceClient
from rich.table import Table

from ...core.context import GraphContext
from ...utils import cache, output
from ...utils.console import console
from ...utils.errors import handle_graph_errors
from ...utils.pagination import GraphPaginator


async def get_user_m365_groups(
    graph_client: GraphServiceClient,
    visibility: str | None = None,
    teams_only: bool = False,
) -> list:
    """Return the current user's M365 Unified groups.

    Includes groups joined via transitive (nested) membership so that indirect
    membership is not missed.

    Args:
        graph_client: Authenticated Graph client.
        visibility:   Optional — 'Private' or 'Public'. None returns both.
        teams_only:   When True, restrict to groups that have a Teams workspace
                      provisioned. Use only when Teams is the actual target;
                      for SharePoint enumeration leave this False so groups
                      without Teams (but with a site) are included.
    """
    logger.debug("Fetching transitive M365 group membership")

    all_groups = await GraphPaginator(
        graph_client.me.transitive_member_of.graph_group
    ).collect()

    # Keep only M365 Unified groups — the only type that owns a SharePoint site.
    # Security and distribution groups are excluded by this filter.
    m365_groups = [
        g for g in all_groups
        if g.group_types and "Unified" in g.group_types
    ]

    if teams_only:
        # "Team" in resourceProvisioningOptions means the group has a Teams
        # workspace. Groups without it may still have SharePoint sites.
        m365_groups = [
            g for g in m365_groups
            if g.resource_provisioning_options
            and "Team" in g.resource_provisioning_options
        ]

    if visibility:
        m365_groups = [g for g in m365_groups if g.visibility == visibility]

    private_count = sum(1 for g in m365_groups if g.visibility == "Private")
    logger.info(
        f"Found {len(m365_groups)} M365 Unified groups "
        f"({private_count} private, {len(m365_groups) - private_count} public) "
        f"out of {len(all_groups)} total group memberships"
    )

    return m365_groups


def add_arguments(parser: argparse.ArgumentParser) -> None:
    parser.add_argument(
        "--visibility",
        choices=["Private", "Public"],
        default=None,
        help="Filter by visibility.",
    )
    parser.add_argument(
        "--teams-only",
        action="store_true",
        help="Only show groups that have a Microsoft Teams workspace provisioned.",
    )


@handle_graph_errors
async def run_with_arguments(context: GraphContext, args: argparse.Namespace) -> int:
    if context.is_app_only:
        logger.error("This module requires delegated authentication (user context).")
        return 1

    groups = await get_user_m365_groups(
        context.graph_client,
        visibility=getattr(args, "visibility", None),
        teams_only=getattr(args, "teams_only", False),
    )

    if not groups:
        logger.warning("No M365 Unified groups found.")
        return 0

    rows = [
        {
            "id": g.id,
            "display_name": g.display_name,
            "mail": g.mail,
            "visibility": g.visibility,
            "teams_provisioned": bool(
                g.resource_provisioning_options
                and "Team" in g.resource_provisioning_options
            ),
        }
        for g in groups
    ]

    cache.save_results(rows, key="sp_groups", identity=context.identity_hash)

    if context.json_output:
        output.print_json(rows)
        return 0

    if context.ndjson_output:
        for row in rows:
            output.print_ndjson_item(row)
        return 0

    table = Table(
        title=f"Your M365 Unified Groups ({len(groups)})",
        show_header=True,
        header_style="bold",
        box=None,
        padding=(0, 1),
    )
    table.add_column("#", justify="right", style="dim")
    table.add_column("Group")
    table.add_column("Email")
    table.add_column("Visibility")
    table.add_column("Teams")

    for idx, (group, row) in enumerate(zip(groups, rows), 1):
        table.add_row(
            str(idx),
            group.display_name or "(unnamed)",
            group.mail or "",
            group.visibility or "",
            "yes" if row["teams_provisioned"] else "",
        )

    console.print(table)
    return 0
