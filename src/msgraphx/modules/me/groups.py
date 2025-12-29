# msgraphx/modules/me/groups.py

# Built-in imports
from typing import TYPE_CHECKING
from datetime import datetime, timezone

# External library imports
from loguru import logger
from msgraphx.utils.errors import handle_graph_errors

if TYPE_CHECKING:
    import argparse
    from msgraphx.core.context import GraphContext


@handle_graph_errors
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


async def run_with_arguments(
    context: "GraphContext", args: "argparse.Namespace"
) -> int:
    logger.info("📊 Microsoft 365 Groups")

    groups = await fetch_user_groups(context)

    # Apply filters
    if args.visibility:
        groups = [g for g in groups if g.visibility == args.visibility]

    if groups:
        # Sort by creation date (newest first)
        groups.sort(
            key=lambda g: g.created_date_time
            or datetime.min.replace(tzinfo=timezone.utc),
            reverse=True,
        )

        for group in groups:
            logger.success(f"{group.display_name}")

            if group.description:
                logger.info(f"Description: {group.description}")

            logger.info(f"ID: {group.id}")

            if group.mail:
                logger.info(f"Email: {group.mail}")

            if group.security_identifier:
                logger.info(f"Security Identifier: {group.security_identifier}")

            # Show boolean properties
            logger.info(f"Security Enabled: {group.security_enabled}")
            logger.info(f"Mail Enabled: {group.mail_enabled}")

            if group.visibility:
                logger.info(f"Visibility: {group.visibility}")

            # Show creation date if available
            if group.created_date_time:
                logger.info(
                    f"Created: {group.created_date_time.strftime('%Y-%m-%d %H:%M:%S UTC')}"
                )

        logger.info(f"Total: {len(groups)} groups")
    else:
        logger.info("📭 No groups found")

    return 0
