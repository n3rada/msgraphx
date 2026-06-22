# msgraphx/modules/onedrive/search.py

# Built-in imports
from __future__ import annotations

import argparse

# External library imports
from loguru import logger
from msgraph.generated.models.drive_item import DriveItem
from msgraph.generated.models.entity_type import EntityType
from msgraph.generated.models.search_content import SearchContent
from msgraph.generated.models.share_point_one_drive_options import SharePointOneDriveOptions

# Local library imports
from ...core import graph_search
from ...core.context import GraphContext
from ...utils import cache, output
from ...utils.console import console
from ...utils.dates import parse_date_string
from ...utils.errors import handle_graph_errors

_ONEDRIVE_ONLY = SharePointOneDriveOptions(include_content=SearchContent.PrivateContent)

HUNT_QUERIES = {
    "credentials": (
        "((filetype:key OR filetype:pem OR filetype:crt OR filetype:cer OR filetype:kdbx "
        "OR filetype:pfx) "
        "OR ((filetype:env OR filetype:cfg OR filetype:yaml OR filetype:yml OR filetype:secret) "
        'AND ("password" OR "passwd" OR "secret" OR "api_key" OR "access_key" OR "client_secret")))'
    ),
    "ssh": (
        "(filetype:pub OR filetype:pem OR filename:id_rsa OR filename:id_ecdsa OR "
        "filename:id_ed25519 OR filename:id_dsa OR filename:authorized_keys OR "
        'filename:known_hosts OR "BEGIN RSA PRIVATE KEY" OR '
        '"BEGIN OPENSSH PRIVATE KEY" OR "BEGIN EC PRIVATE KEY" OR "BEGIN PRIVATE KEY")'
    ),
    "office": (
        "filetype:doc OR filetype:docx OR filetype:xls OR filetype:xlsx OR "
        "filetype:ppt OR filetype:pptx OR filetype:pdf"
    ),
    "scripts": (
        "(filetype:ps1 OR filetype:sh OR filetype:bat OR filetype:cmd OR "
        "filetype:py OR filetype:rb OR filetype:pl OR filetype:ts)"
    ),
    "configs": (
        "((filetype:conf OR filetype:ini OR filetype:env OR filetype:yaml OR filetype:yml) "
        'AND ("password" OR "secret" OR "token" OR "credentials"))'
    ),
}


async def fetch(
    context: GraphContext,
    query: str = "*",
    after: str | None = None,
    before: str | None = None,
    region: str | None = None,
) -> list[dict]:
    """Search personal OneDrive drives across the tenant.

    With delegated auth, searches the current user's OneDrive.
    With app-only auth and a region, searches all personal drives in the tenant.
    """
    search_query = query

    filters = []
    if before:
        filters.append(f"created<={before}")
    if after:
        filters.append(f"created>={after}")
    if filters:
        search_query += " " + " ".join(filters)

    options = graph_search.SearchOptions(
        query_string=search_query,
        page_size=500,
        region=region,
        share_point_one_drive_options=_ONEDRIVE_ONLY,
    )

    items: list[dict] = []
    async for _item in graph_search.search_entities(
        context.graph_client,
        entity_types=[EntityType.DriveItem],
        options=options,
    ):
        drive_item = _item if isinstance(_item, DriveItem) else None
        if drive_item is None:
            continue

        author = (
            drive_item.created_by.user.display_name
            if drive_item.created_by and drive_item.created_by.user
            else "?"
        )
        size_bytes = drive_item.size or 0
        if size_bytes >= 1_048_576:
            size_str = f"{size_bytes / 1_048_576:.1f} MB"
        elif size_bytes >= 1024:
            size_str = f"{size_bytes / 1024:.1f} KB"
        else:
            size_str = f"{size_bytes} B"

        items.append({
            "drive_id": (
                drive_item.parent_reference.drive_id if drive_item.parent_reference else None
            ),
            "item_id": drive_item.id,
            "name": drive_item.name,
            "size": drive_item.size,
            "size_str": size_str,
            "web_url": drive_item.web_url,
            "created_datetime": (
                drive_item.created_date_time.isoformat() if drive_item.created_date_time else None
            ),
            "last_modified_datetime": (
                drive_item.last_modified_date_time.isoformat()
                if drive_item.last_modified_date_time
                else None
            ),
            "created": (
                drive_item.created_date_time.strftime("%Y-%m-%d")
                if drive_item.created_date_time
                else ""
            ),
            "author": author,
            "created_by_email": (
                drive_item.created_by.user.additional_data.get("email")
                if drive_item.created_by and drive_item.created_by.user
                else None
            ),
            "created_by_id": (
                drive_item.created_by.user.id
                if drive_item.created_by and drive_item.created_by.user
                else None
            ),
            "last_modified_by": (
                drive_item.last_modified_by.user.display_name
                if drive_item.last_modified_by and drive_item.last_modified_by.user
                else None
            ),
            "parent_path": (
                drive_item.parent_reference.path if drive_item.parent_reference else None
            ),
            "site_id": (
                drive_item.parent_reference.site_id if drive_item.parent_reference else None
            ),
            "mime_type": (drive_item.file.mime_type if drive_item.file else None),
            "is_folder": drive_item.folder is not None,
            "download_url": (
                drive_item.additional_data.get("@microsoft.graph.downloadUrl")
                if drive_item.additional_data
                else None
            ),
        })

    return items


def add_arguments(parser: argparse.ArgumentParser) -> None:
    parser.set_defaults(uses_time_bounds=True)
    parser.add_argument(
        "query",
        nargs="?",
        default="*",
        help="Search query (KQL). Defaults to '*'.",
    )
    parser.add_argument(
        "-f",
        "--filetype",
        type=str,
        help="Filter by file extension (e.g. pdf, docx).",
    )
    parser.add_argument(
        "--hunt",
        choices=HUNT_QUERIES.keys(),
        help="Predefined hunt query (credentials, ssh, office, scripts, configs).",
    )


@handle_graph_errors
async def run_with_arguments(context: GraphContext, args: argparse.Namespace) -> int:
    search_query = args.query

    if args.filetype:
        search_query = f"filetype:{args.filetype} {search_query}".strip()
    elif args.hunt:
        search_query = HUNT_QUERIES[args.hunt]
        logger.info("Hunt mode")

    after_iso: str | None = None
    before_iso: str | None = None

    if args.before:
        try:
            before_iso = parse_date_string(args.before)
        except ValueError as e:
            logger.error(str(e))
            return 1

    if args.after:
        try:
            after_iso = parse_date_string(args.after)
        except ValueError as e:
            logger.error(str(e))
            return 1

    region = args.region if getattr(args, "is_app_only", False) else None

    logger.info(f"Search query: {search_query}")

    items = await fetch(context, query=search_query, after=after_iso, before=before_iso, region=region)

    if not items:
        logger.info("No results found.")
        if context.json_output:
            output.print_json([])
        return 0

    cache.save_results(items, key="onedrive")

    if context.json_output:
        output.print_json(items)
        return 0

    if context.ndjson_output:
        for item in items:
            output.print_ndjson_item(item)
        return 0

    console.print(f"[bold]OneDrive search results[/bold] ({len(items)})")
    console.rule()
    for count, item in enumerate(items, 1):
        console.print(
            f"  [dim]{count:>4}.[/dim]  {item['name']}  "
            f"[dim]{item['author']}[/dim]  [cyan]{item['size_str']}[/cyan]  [dim]{item['created']}[/dim]"
        )

    return 0
