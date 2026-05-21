# msgraphx/modules/sharepoint/sites.py
#
# List SharePoint sites the current user has access to via M365 group membership.
# Fetches the user's groups, filters for M365 (Unified) groups, then resolves
# each group's root SharePoint site URL in parallel.
#
# Also shows sites the user is directly following (GET /me/followedSites).
#
# Required delegated permissions: Sites.Read.All, Group.Read.All

# Built-in imports
from __future__ import annotations

import argparse
import asyncio

# External library imports
from loguru import logger
from rich.table import Table

# Local library imports
from .groups import get_user_m365_groups
from ...core.context import GraphContext
from ...utils.console import console
from ...utils.errors import handle_graph_errors


async def _fetch_group_site(context: "GraphContext", group_id: str) -> tuple | None:
    """Fetch root SharePoint site for a group. Returns (site_name, web_url) or None."""
    try:
        site = (
            await context.graph_client.groups.by_group_id(group_id)
            .sites.by_site_id("root")
            .get()
        )
        if site:
            return (
                site.display_name or site.name or "(unnamed site)",
                site.web_url or "",
            )
    except Exception as exc:
        logger.debug(
            f"Failed to fetch site for group {group_id}: {type(exc).__name__}: {exc}"
        )
        raise  # Let gather collect the exception
    return None


@handle_graph_errors
async def run_with_arguments(
    context: "GraphContext", args: "argparse.Namespace"
) -> int:
    if context.is_app_only:
        logger.error("This module requires delegated authentication (user context).")
        return 1

    show_all = getattr(args, "all_visibility", False)


    # -------------------------------------------------------------------------
    # 1. Fetch user's M365 group memberships
    # -------------------------------------------------------------------------
    m365_groups = await get_user_m365_groups(context.graph_client)
    private_groups = [g for g in m365_groups if g.visibility == "Private"]
    public_groups = [g for g in m365_groups if g.visibility != "Private"]

    # -------------------------------------------------------------------------
    # 2. Resolve SharePoint site URLs for target groups (in parallel)
    # -------------------------------------------------------------------------
    target_groups = m365_groups if show_all else private_groups
    logger.info(f"Resolving SharePoint sites for {len(target_groups)} groups")

    # Fetch all group sites concurrently
    tasks = [_fetch_group_site(context, g.id) for g in target_groups]
    results = await asyncio.gather(*tasks, return_exceptions=True)

    private_rows: list[tuple] = []
    public_rows: list[tuple] = []
    errors = 0

    for group, result in zip(target_groups, results):
        if isinstance(result, Exception):
            errors += 1
            if errors == 1:
                logger.warning(
                    f"Failed to fetch site for '{group.display_name}': {result}"
                )
            continue
        if result is None:
            logger.debug(f"  No site for group: {group.display_name}")
            continue

        site_name, web_url = result
        visibility = group.visibility or "Public"
        row = (group.display_name or "(unnamed)", site_name, web_url)

        logger.debug(f"  {visibility}: {group.display_name} -> {web_url}")

        if visibility == "Private":
            private_rows.append(row)
        else:
            public_rows.append(row)

    if errors:
        logger.warning(f"{errors}/{len(target_groups)} group site lookups failed")

    # -------------------------------------------------------------------------
    # 3. Followed sites
    # -------------------------------------------------------------------------
    logger.info("Fetching followed sites")
    followed_rows: list[tuple] = []
    try:
        followed_resp = await context.graph_client.me.followed_sites.get()
        followed = followed_resp.value if followed_resp and followed_resp.value else []
        for site in followed:
            followed_rows.append((site.display_name or "", site.web_url or ""))
            logger.debug(f"  Followed: {site.display_name} | {site.web_url}")
    except Exception as exc:
        logger.debug(f"Failed to fetch followed sites: {exc}")

    # -------------------------------------------------------------------------
    # Display
    # -------------------------------------------------------------------------
    def _print_table(title: str, columns: list[str], rows: list[tuple]) -> None:
        if not rows:
            return
        table = Table(
            title=title,
            show_header=True,
            header_style="bold",
            box=None,
            padding=(0, 1),
        )
        for col in columns:
            table.add_column(col)
        for row in rows:
            table.add_row(*row)
        console.print(table)
        console.print()

    _print_table(
        f"Private sites ({len(private_rows)}) - via group membership",
        ["Group", "Site", "URL"],
        private_rows,
    )

    if show_all:
        _print_table(
            f"Public sites ({len(public_rows)}) - via group membership",
            ["Group", "Site", "URL"],
            public_rows,
        )

    _print_table(
        f"Followed sites ({len(followed_rows)})",
        ["Site", "URL"],
        followed_rows,
    )

    total = (
        len(private_rows) + (len(public_rows) if show_all else 0) + len(followed_rows)
    )
    if total == 0:
        logger.info("No sites found")
    else:
        logger.success(f"{total} site(s) listed")

    return 0


def add_arguments(parser: "argparse.ArgumentParser") -> None:
    parser.add_argument(
        "--public",
        dest="all_visibility",
        action="store_true",
        help="Also show public group sites (default: private only).",
    )
