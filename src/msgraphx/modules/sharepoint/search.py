# msgraphx/modules/sharepoint/search.py

# Built-in imports
from typing import TYPE_CHECKING
import argparse

# External library imports
from loguru import logger
from msgraph import GraphServiceClient
from msgraph.generated.models.entity_type import EntityType

# Local library imports
from msgraphx.core import graph_search
from msgraphx.utils.dates import parse_date_string


if TYPE_CHECKING:
    import argparse
    from msgraphx.core.context import GraphContext

HUNT_QUERIES = {
    "scripts": (
        "filetype:ps1 OR filetype:sh OR filetype:bat OR filetype:cmd OR "
        "filetype:py OR filetype:rb OR filetype:pl OR filetype:java OR "
        "filetype:ts OR filetype:cs OR filetype:cpp OR filetype:vbs"
    ),
    "credentials": (
        "filetype:key OR filetype:pem OR filetype:crt OR filetype:cer OR filetype:env OR "
        "filetype:kdbx OR filetype:cfg OR filetype:yaml OR "
        "filetype:yml OR filetype:secret OR filetype:vault OR filetype:pfx"
    ),
    "office": (
        "filetype:doc OR filetype:docx OR filetype:rtf OR filetype:odt OR "
        "filetype:xls OR filetype:xlsx OR filetype:csv OR filetype:ods OR "
        "filetype:ppt OR filetype:pptx OR filetype:pdf OR filetype:msg OR filetype:eml"
    ),
    "backups": (
        "filetype:bak OR filetype:zip OR filetype:rar OR filetype:7z OR "
        "filetype:gz OR filetype:tar OR filetype:tgz OR filetype:tar.gz OR "
        "filetype:db OR filetype:sqlite OR filetype:mdb OR filetype:accdb OR "
        "filetype:vhd OR filetype:vmdk OR filetype:ova OR filetype:old"
    ),
    "configs": (
        "filetype:conf OR filetype:config OR filetype:ini OR "
        "filetype:yaml OR filetype:yml OR filetype:psd1 OR filetype:reg"
    ),
    "infra": (
        "filetype:dockerfile OR filetype:compose OR filetype:tf OR "
        "filetype:terraform"
    ),
    "network": ("filetype:pcap OR filetype:cap OR filetype:har"),
}


async def get_user_sharepoint_groups(
    graph_client: "GraphServiceClient", visibility: str = None
) -> list:
    """
    Get current user's Microsoft 365 groups (which have SharePoint sites).

    Args:
        graph_client: Authenticated Graph client
        visibility: Optional filter - 'Private' or 'Public'

    Returns:
        List of group objects with SharePoint sites
    """
    logger.info("📋 Fetching user's Microsoft 365 groups...")

    try:
        # Build filter for M365 groups
        filters = ["groupTypes/any(c:c eq 'Unified')"]
        if visibility:
            filters.append(f"visibility eq '{visibility}'")

        # Use direct query without explicit builder import
        result = await graph_client.me.transitive_member_of.graph_group.get()

        if result and result.value:
            # Apply client-side filters
            sp_groups = result.value

            # Filter by visibility if specified
            if visibility:
                sp_groups = [g for g in sp_groups if g.visibility == visibility]

            # Filter for groups with Team/SharePoint
            sp_groups = [
                g
                for g in result.value
                if g.resource_provisioning_options
                and "Team" in g.resource_provisioning_options
            ]
            logger.success(
                f"✅ Found {len(sp_groups)} Microsoft 365 groups with SharePoint sites"
            )
            return sp_groups

        logger.info("📭 No Microsoft 365 groups found")
        return []

    except Exception as e:
        logger.error(f"❌ Failed to fetch user groups: {e}")
        return []


def add_arguments(parser: "argparse.ArgumentParser"):
    parser.add_argument(
        "query",
        nargs="?",
        default="*",
        help="Microsoft Search query string (e.g., filetype:pdf, name:\"report\"). Defaults to '*'.",
    )

    parser.add_argument(
        "--site",
        type=str,
        default=None,
        help="SharePoint site to process. Can be the full site URL or the site ID.",
    )

    parser.add_argument(
        "-f",
        "--filetype",
        type=str,
        help="Shortcut to set --query='filetype:<ext>' (e.g., pdf, docx)",
    )

    parser.add_argument(
        "--hunt",
        choices=HUNT_QUERIES.keys(),
        help="Shortcut to set --query for specific file types (scripts, credentials, office, backups)",
    )

    parser.add_argument(
        "--my-groups",
        action="store_true",
        help="Search only in current user's Microsoft 365 groups (SharePoint sites).",
    )

    parser.add_argument(
        "--list-groups",
        action="store_true",
        help="List current user's Microsoft 365 groups and exit (no search).",
    )

    parser.add_argument(
        "--visibility",
        type=str,
        choices=["Private", "Public"],
        default=None,
        help="Filter groups by visibility (use with --my-groups or --list-groups).",
    )


async def run_with_arguments(
    context: "GraphContext", args: "argparse.Namespace"
) -> int:

    # Handle --list-groups command
    if args.list_groups:
        groups = await get_user_sharepoint_groups(context.graph_client, args.visibility)

        if groups:
            logger.info("📊 Your Microsoft 365 Groups with SharePoint:")
            for group in groups:
                visibility = f"🔒 {group.visibility}" if group.visibility else ""
                logger.success(f"  👥 {group.displayName} {visibility}")
                if group.mail:
                    logger.info(f"     Email: {group.mail}")
                logger.info(f"     ID: {group.id}")

        return 0

    search_query = args.query

    if args.filetype:
        # Add filetype filter to the query
        search_query = f"filetype:{args.filetype} {search_query}".strip()

    elif args.hunt:
        # Use predefined hunt query
        search_query = HUNT_QUERIES[args.hunt]
        logger.info("🎯 Hunt mode")

    filters = []

    if args.before:
        try:
            iso_date = parse_date_string(args.before)
            filters.append(f"created<={iso_date}")
        except ValueError as e:
            logger.error(str(e))
            return 1

    if args.after:
        try:
            iso_date = parse_date_string(args.after)
            filters.append(f"created>={iso_date}")
        except ValueError as e:
            logger.error(str(e))
            return 1

    if filters:
        search_query += " " + " ".join(filters)
        logger.info(f"📅 Date filters applied: {' and '.join(filters)}")

    logger.info(f"🔍 Search query: {search_query}")

    # Region is required for application permissions, optional for delegated
    region = args.region if getattr(args, "is_app_only", False) else None

    # Handle --my-groups: search only in user's groups
    group_ids = None
    if args.my_groups:
        groups = await get_user_sharepoint_groups(context.graph_client, args.visibility)
        if not groups:
            logger.warning(
                "⚠️ No Microsoft 365 groups found, no results will be returned"
            )
            return 0

        group_ids = [g.id for g in groups]
        logger.info(f"🔒 Scoping search to {len(group_ids)} user groups")

    # Get drive_id if provided to scope search
    drive_id = getattr(args, "drive_id", None)
    if drive_id:
        logger.info(f"🔒 Scoping search to Drive ID: {drive_id}")

    # Build search query with group filter if needed
    if group_ids:
        # Add filter to search only in these groups' drives
        group_filter = " OR ".join([f"GroupId:{gid}" for gid in group_ids])
        search_query = f"({search_query}) AND ({group_filter})"
        logger.debug(f"📝 Modified query: {search_query}")

    search_options = graph_search.SearchOptions(
        query_string=search_query,
        sort_by="createdDateTime",
        descending=True,
        page_size=500,
        region=region,
        drive_id=drive_id,
    )

    count = 0
    async for drive_item in graph_search.search_entities(
        context.graph_client,
        entity_types=[EntityType.DriveItem],
        options=search_options,
    ):
        # Process each DriveItem
        logger.info(
            f"📄 {drive_item.name} (Created by: {drive_item.created_by.user.display_name})"
        )
        count += 1

    logger.info(f"📊 Total files found: {count}")
    return 0
