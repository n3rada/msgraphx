# Built-in imports
from typing import TYPE_CHECKING
import argparse

# External library imports
from loguru import logger
from msgraph import GraphServiceClient
from msgraph.generated.models.entity_type import EntityType

# Local library imports
from msgraphx.core import search
from msgraphx.utils.dates import parse_date_string


if TYPE_CHECKING:
    import argparse
    from msgraph import GraphServiceClient

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


async def run_with_arguments(
    graph_client: "GraphServiceClient", args: "argparse.Namespace"
) -> int:

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

    # Get drive_id if provided to scope search
    drive_id = getattr(args, "drive_id", None)
    if drive_id:
        logger.info(f"🔒 Scoping search to Drive ID: {drive_id}")

    search_options = search.SearchOptions(
        query_string=search_query,
        sort_by="createdDateTime",
        descending=True,
        page_size=500,
        region=region,
        drive_id=drive_id,
    )

    async for drive_item in search.search_entities(
        graph_client,
        entity_types=[EntityType.DriveItem],
        options=search_options,
    ):
        # Process each DriveItem
        logger.info(
            f"📄 {drive_item.name} (Created by: {drive_item.created_by.user.display_name})"
        )

    return 0
