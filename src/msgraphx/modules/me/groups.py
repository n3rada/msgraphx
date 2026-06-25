# msgraphx/modules/me/groups.py

# Built-in imports
from __future__ import annotations

import argparse
from datetime import datetime, timezone

# External library imports
from loguru import logger

# Local library imports
from ...core.context import GraphContext
from ...utils import output
from ...utils.errors import handle_graph_errors
from ...utils.roles import require_scopes


async def fetch(context: GraphContext, visibility: str | None = None) -> list[dict]:
    """Return groups the current user belongs to (transitive) as plain dicts.

    Raises on API error. Callers are responsible for handling exceptions.
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


def add_arguments(parser: "argparse.ArgumentParser"):
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
    context: "GraphContext", args: "argparse.Namespace"
) -> int:
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
