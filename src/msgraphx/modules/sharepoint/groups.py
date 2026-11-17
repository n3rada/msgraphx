# msgraphx/modules/sharepoint/groups.py
#
# Helper for fetching the current user's M365 Unified groups. Used to scope
# SharePoint search and site enumeration, and by `groups list --mine`.
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

from __future__ import annotations

from loguru import logger
from msgraph.generated.models.group import Group
from msgraph.graph_service_client import GraphServiceClient

from ...utils.pagination import GraphPaginator

async def get_user_m365_groups(
    graph_client: GraphServiceClient,
    visibility: str | None = None,
    teams_only: bool = False,
) -> list[Group]:
    """Return the signed-in user's M365 Unified groups.

    Includes groups joined via transitive (nested) membership so that indirect
    membership is not missed. Requires a delegated token (`/me`).

    Args:
        graph_client: Authenticated Graph client.
        visibility: "Private" or "Public" to keep only that visibility. None
            returns both.
        teams_only: When True, keep only groups with a Teams workspace
            provisioned. Use only when Teams is the actual target; for
            SharePoint enumeration leave this False so groups without Teams
            (but with a site) are included.

    Returns:
        The matching Unified groups, in server order.
    """
    logger.debug("Fetching transitive M365 group membership")

    all_groups: list[Group] = await GraphPaginator(
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
