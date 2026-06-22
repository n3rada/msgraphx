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
from ...utils import output
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
        logger.error(f"Failed to fetch site for group {group_id}: {exc}")
        raise
    return None


async def fetch(context: GraphContext, show_public: bool = False) -> dict:
    """Return SharePoint sites accessible to the current user as plain dicts.

    Returns a dict with keys 'private_sites', 'public_sites', 'followed_sites'.
    Raises on API error — callers are responsible for handling exceptions.
    """
    m365_groups = await get_user_m365_groups(context.graph_client)
    private_groups = [g for g in m365_groups if g.visibility == "Private"]
    public_groups = [g for g in m365_groups if g.visibility != "Private"]

    target_groups = m365_groups if show_public else private_groups
    tasks = [_fetch_group_site(context, g.id) for g in target_groups]
    results = await asyncio.gather(*tasks, return_exceptions=True)

    private_sites: list[dict] = []
    public_sites: list[dict] = []

    for group, result in zip(target_groups, results):
        if isinstance(result, Exception):
            continue
        if result is None:
            continue
        site_name, web_url = result
        visibility = group.visibility or "Public"
        entry = {"group": group.display_name or "(unnamed)", "site": site_name, "url": web_url}
        if visibility == "Private":
            private_sites.append(entry)
        else:
            public_sites.append(entry)

    followed_sites: list[dict] = []
    followed_resp = await context.graph_client.me.followed_sites.get()
    followed = followed_resp.value if followed_resp and followed_resp.value else []
    for site in followed:
        followed_sites.append({"site": site.display_name or "", "url": site.web_url or ""})

    return {
        "private_sites": private_sites,
        "public_sites": public_sites if show_public else [],
        "followed_sites": followed_sites,
    }


@handle_graph_errors
async def run_with_arguments(
    context: "GraphContext", args: "argparse.Namespace"
) -> int:
    if context.is_app_only:
        logger.error("This module requires delegated authentication (user context).")
        return 1

    show_all = getattr(args, "all_visibility", False)
    logger.info("Fetching SharePoint sites")

    data = await fetch(context, show_public=show_all)

    private_rows = [(s["group"], s["site"], s["url"]) for s in data["private_sites"]]
    public_rows = [(s["group"], s["site"], s["url"]) for s in data["public_sites"]]
    followed_rows = [(s["site"], s["url"]) for s in data["followed_sites"]]
    total = len(private_rows) + len(public_rows) + len(followed_rows)

    if context.json_output:
        output.print_json(data)
        return 0

    if context.ndjson_output:
        output.print_ndjson_item(data)
        return 0

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
