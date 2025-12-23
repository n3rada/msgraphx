# msgraphx/modules/me/groups.py

# Built-in imports
from typing import TYPE_CHECKING
from datetime import datetime, timezone

# External library imports
from loguru import logger

if TYPE_CHECKING:
    import argparse
    from msgraphx.core.context import GraphContext


async def fetch_user_groups(context: "GraphContext") -> list:
    """
    Fetch all groups the current user is a member of (transitive).

    Args:
        context: Graph context with authenticated client

    Returns:
        List of group objects
    """
    try:
        result = await context.graph_client.me.transitive_member_of.graph_group.get()
        return result.value if result and result.value else []
    except Exception as e:
        logger.error(f"❌ Failed to fetch user groups: {e}")
        return []


def add_arguments(parser: "argparse.ArgumentParser"):
    parser.add_argument(
        "--visibility",
        type=str,
        choices=["Private", "Public"],
        default=None,
        help="Filter groups by visibility.",
    )

    parser.add_argument(
        "--with-sharepoint",
        action="store_true",
        help="Only show groups with SharePoint sites.",
    )


async def run_with_arguments(
    context: "GraphContext", args: "argparse.Namespace"
) -> int:
    logger.info("📊 Microsoft 365 Groups")

    groups = await fetch_user_groups(context)

    # Apply filters
    if args.visibility:
        groups = [g for g in groups if g.visibility == args.visibility]

    if args.with_sharepoint:
        groups = [
            g
            for g in groups
            if g.additional_data
            and g.additional_data.get("resourceProvisioningOptions")
            and "Team" in g.additional_data.get("resourceProvisioningOptions", [])
        ]

    if groups:
        # Sort by creation date (newest first)
        groups.sort(
            key=lambda g: g.created_date_time
            or datetime.min.replace(tzinfo=timezone.utc),
            reverse=True,
        )

        for group in groups:
            # Check if group has SharePoint/Teams
            has_sharepoint = (
                group.additional_data
                and group.additional_data.get("resourceProvisioningOptions")
                and "Team"
                in group.additional_data.get("resourceProvisioningOptions", [])
            )

            sp_indicator = "📁 SharePoint" if has_sharepoint else ""
            visibility_indicator = f"🔒 {group.visibility}" if group.visibility else ""
            security_indicator = "🔐 Security" if group.security_enabled else ""

            # Build indicator string
            indicators = " ".join(
                filter(None, [visibility_indicator, security_indicator, sp_indicator])
            )

            logger.success(f"{group.display_name} {indicators}")

            if group.description:
                logger.info(f"Description: {group.description}")
            if group.mail:
                logger.info(f"Email: {group.mail}")
            logger.info(f"ID: {group.id}")
            if group.security_identifier:
                logger.info(f"Security Identifier: {group.security_identifier}")

            # Show creation date if available
            if group.created_date_time:
                logger.info(
                    f"Created: {group.created_date_time.strftime('%Y-%m-%d %H:%M:%S UTC')}"
                )

        logger.info(f"Total: {len(groups)} groups")
    else:
        logger.info("📭 No groups found")

    return 0
