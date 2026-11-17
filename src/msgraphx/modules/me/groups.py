# msgraphx/modules/me/groups.py

# Built-in imports
from __future__ import annotations

import argparse
from datetime import datetime, timezone
from typing import Any

# External library imports
from loguru import logger

# Local library imports
from ...core.context import GraphContext
from ...utils import output
from ...utils.errors import handle_graph_errors
from ...utils.roles import require_scopes


async def fetch(
    context: GraphContext, visibility: str | None = None
) -> list[dict[str, Any]]:
    """Return the groups the current user transitively belongs to.

    Only the first page of `/me/transitiveMemberOf/microsoft.graph.group`
    is read. Groups are sorted by creation time, newest first.

    Args:
        context: Shared Graph runtime context (delegated auth).
        visibility: Keep only groups with this visibility (`Private` or
            `Public`). `None` keeps all groups.

    Returns:
        One dict per group with the keys `id`, `display_name`,
        `description`, `mail`, `security_identifier`, `security_enabled`,
        `mail_enabled`, `visibility`, and `created_date_time` (ISO 8601 or
        `None`).

    Raises:
        ODataError: If the Graph API rejects the request.
    """
    result = await context.graph_client.me.transitive_member_of.graph_group.get()
    groups = result.value if result and result.value else []

    if visibility:
        groups = [g for g in groups if g.visibility == visibility]

    groups.sort(
        key=lambda g: g.created_date_time or datetime.min.replace(tzinfo=timezone.utc),
        reverse=True,
    )

    return [
        {
            "id": g.id,
            "display_name": g.display_name,
            "description": g.description,
            "mail": g.mail,
            "security_identifier": g.security_identifier,
            "security_enabled": g.security_enabled,
            "mail_enabled": g.mail_enabled,
            "visibility": g.visibility,
            "created_date_time": g.created_date_time.isoformat() if g.created_date_time else None,
        }
        for g in groups
    ]


def add_arguments(parser: argparse.ArgumentParser) -> None:
    """Register the `me groups` options on the parser.

    Args:
        parser: Parser for the `me groups` subcommand.
    """
    parser.add_argument(
        "--visibility",
        type=str,
        choices=["Private", "Public"],
        default=None,
        help="Filter groups by visibility.",
    )


@handle_graph_errors
@require_scopes("Group.Read.All")
async def run_with_arguments(
    context: GraphContext, args: argparse.Namespace
) -> int:
    """List the current user's group memberships.

    Logs each group's details, or prints JSON/NDJSON when the matching
    output flag is set.

    Args:
        context: Shared Graph runtime context.
        args: Parsed CLI arguments (`visibility`).

    Returns:
        0 on success, including when no groups are found.

    Raises:
        AuthenticationError: If the Graph token is invalid or expired.
        ODataError: If the Graph API rejects the request.
    """
    logger.info("Microsoft 365 Groups")

    rows = await fetch(context, visibility=args.visibility)

    if not rows:
        logger.info("No groups found")
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

    for row in rows:
        logger.success(f"{row['display_name']}")

        if row["description"]:
            logger.info(f"Description: {row['description']}")

        logger.info(f"ID: {row['id']}")

        if row["mail"]:
            logger.info(f"Email: {row['mail']}")

        if row["security_identifier"]:
            logger.info(f"Security Identifier: {row['security_identifier']}")

        logger.info(f"Security Enabled: {row['security_enabled']}")
        logger.info(f"Mail Enabled: {row['mail_enabled']}")

        if row["visibility"]:
            logger.info(f"Visibility: {row['visibility']}")

        if row["created_date_time"]:
            logger.info(f"Created: {row['created_date_time']}")

    logger.info(f"Total: {len(rows)} groups")
    return 0
