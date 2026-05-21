# msgraphx/modules/sharepoint/groups.py
#
# Shared helper and CLI subcommand for fetching the current user's M365 groups.
# The helper `get_user_m365_groups` is reused by search.py and sites.py.
# CLI: msgraphx sp groups

from __future__ import annotations

import argparse

from loguru import logger
from msgraph.graph_service_client import GraphServiceClient
from rich.table import Table

from ...utils.console import console

from ...core.context import GraphContext
from ...utils.errors import handle_graph_errors
from ...utils.pagination import GraphPaginator


# ---------------------------------------------------------------------------
# Shared helper (importable by other modules)
# ---------------------------------------------------------------------------


async def get_user_m365_groups(
    graph_client: "GraphServiceClient",
    visibility: str | None = None,
    teams_only: bool = False,
) -> list:
    """
    Fetch the current user's M365 (Unified) groups via transitiveMemberOf.

    Paginates through all results. Optionally filters by visibility
    and/or whether Teams is provisioned on the group.

    Args:
        graph_client: Authenticated Graph client
        visibility: Optional filter - 'Private' or 'Public'
        teams_only: If True, only return groups with Teams provisioned

    Returns:
        List of Group objects
    """
    logger.info("Fetching user's M365 groups")

    groups = await GraphPaginator(
        graph_client.me.transitive_member_of.graph_group
    ).collect()

    # Only Unified (M365) groups have SharePoint sites
    m365_groups = [g for g in groups if g.group_types and "Unified" in g.group_types]

    if teams_only:
        m365_groups = [
            g
            for g in m365_groups
            if g.resource_provisioning_options
            and "Team" in g.resource_provisioning_options
        ]

    if visibility:
        m365_groups = [g for g in m365_groups if g.visibility == visibility]

    private_count = sum(1 for g in m365_groups if g.visibility == "Private")
    logger.info(
        f"Found {len(m365_groups)} M365 groups ({private_count} private, "
        f"{len(m365_groups) - private_count} public) out of {len(groups)} total"
    )

    return m365_groups


# ---------------------------------------------------------------------------
# CLI subcommand
# ---------------------------------------------------------------------------


def add_arguments(parser: "argparse.ArgumentParser") -> None:
    parser.add_argument(
        "--visibility",
        type=str,
        choices=["Private", "Public"],
        default=None,
        help="Filter groups by visibility.",
    )
    parser.add_argument(
        "--teams-only",
        action="store_true",
        help="Only show groups with Microsoft Teams provisioned.",
    )


@handle_graph_errors
async def run_with_arguments(
    context: "GraphContext", args: "argparse.Namespace"
) -> int:
    if context.is_app_only:
        logger.error("This module requires delegated authentication (user context).")
        return 1

    groups = await get_user_m365_groups(
        context.graph_client,
        visibility=getattr(args, "visibility", None),
        teams_only=getattr(args, "teams_only", False),
    )

    if not groups:
        logger.warning("No M365 groups found.")
        return 0

    table = Table(
        title=f"Your Microsoft 365 Groups ({len(groups)})",
        show_header=True,
        header_style="bold",
        box=None,
        padding=(0, 1),
    )
    table.add_column("#", justify="right", style="dim")
    table.add_column("Group")
    table.add_column("Email")
    table.add_column("Visibility")

    for idx, group in enumerate(groups, 1):
        table.add_row(
            str(idx),
            group.display_name or "(unnamed)",
            group.mail or "",
            group.visibility or "",
        )

    console.print(table)
    return 0
