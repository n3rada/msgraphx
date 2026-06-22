# msgraphx/modules/sharepoint/search.py

# Built-in imports
from __future__ import annotations

import argparse
import json
from pathlib import Path

# External library imports
from loguru import logger
from msgraph.generated.models.drive_item import DriveItem
from msgraph.generated.models.entity_type import EntityType
from rich.table import Table

# Local library imports
from .groups import get_user_m365_groups
from ...core import graph_search
from ...core.context import GraphContext
from ...utils import cache, output
from ...utils.console import console
from ...utils.dates import parse_date_string
from ...utils.errors import handle_graph_errors

HUNT_QUERIES = {
    "scripts": (
        # Match either a known script filetype OR a high-confidence script marker
        "((filetype:ps1 OR filetype:sh OR filetype:bat OR filetype:cmd OR "
        "filetype:py OR filetype:rb OR filetype:pl OR filetype:ts) "
        'OR ("#!" OR "#!/" OR "param(" OR "function " OR "Import-Module"))'
    ),
    "credentials": (
        # High-confidence secret containers OR config-like files that include secret markers
        "((filetype:key OR filetype:pem OR filetype:crt OR filetype:cer OR filetype:kdbx "
        "OR filetype:pfx) "
        "OR ((filetype:env OR filetype:cfg OR filetype:yaml OR filetype:yml OR filetype:secret) "
        'AND ("password" OR "passwd" OR "secret" OR "api_key" OR "access_key" OR "client_secret")))'
    ),
    "ssh": (
        # Filename patterns for private keys or explicit key block headers
        "(filetype:pub OR filetype:pem OR filename:id_rsa OR filename:id_ecdsa OR "
        "filename:id_ed25519 OR filename:id_dsa OR filename:authorized_keys OR "
        'filename:known_hosts OR filename:.ssh OR "BEGIN RSA PRIVATE KEY" OR '
        '"BEGIN OPENSSH PRIVATE KEY" OR "BEGIN DSA PRIVATE KEY" OR '
        '"BEGIN EC PRIVATE KEY" OR "BEGIN PRIVATE KEY" OR "ssh-rsa" OR '
        '"ssh-ed25519" OR "ecdsa-sha2")'
    ),
    "mssql": (
        # Keep focused connection-string and DDL checks to avoid broad matches
        "( (filetype:mdf OR filetype:ldf OR filetype:udl OR filetype:sql) "
        'OR ("Data Source=" AND "Initial Catalog=") '
        'OR ("connection string" AND ("User ID=" OR "Uid=" OR "Password=")) '
        'OR (filename:web.config AND "connectionStrings") '
        'OR (filename:appsettings.json AND "ConnectionStrings") '
        'OR (filetype:sql AND ("CREATE TABLE" OR "INSERT INTO" OR "ALTER TABLE")) )'
    ),
    "office": (
        # Office document filetypes only (keep as explicit filetypes)
        "filetype:doc OR filetype:docx OR filetype:rtf OR filetype:odt OR "
        "filetype:xls OR filetype:xlsx OR filetype:csv OR filetype:ods OR "
        "filetype:ppt OR filetype:pptx OR filetype:pdf OR filetype:msg OR filetype:eml"
    ),
    "backups": (
        # Prefer explicit backup extensions or compressed files that also mention backup/dump
        "(filetype:bak OR filetype:vhd OR filetype:vmdk OR filetype:ova) "
        "OR ((filetype:zip OR filetype:rar OR filetype:7z OR filetype:gz OR filetype:tar) "
        'AND ("backup" OR "dump" OR "dbbackup" OR "backup_")) '
        'OR (filetype:sql AND "backup")'
    ),
    "configs": (
        # Configuration files that explicitly contain credential/config markers
        "((filetype:conf OR filetype:ini OR filetype:psd1 OR filetype:reg) "
        "OR filename:appsettings.json OR filename:web.config OR filename:settings.json) "
        'AND ("connection" OR "password" OR "secret" OR "token" OR "credentials")'
    ),
    "infra": (
        # Common infra-as-code filetypes; require provider/resource keywords for TF/YAML when possible
        "(filetype:dockerfile OR filename:Dockerfile OR filetype:compose OR "
        '(filetype:tf OR filename:terraform) AND ("resource" OR "provider" OR "module"))'
    ),
    "network": ("filetype:pcap OR filetype:cap OR filetype:har"),
}


async def fetch(
    context: GraphContext,
    query: str = "*",
    sort_by: str = "createdDateTime",
    descending: bool = True,
    page_size: int = 500,
    region: str | None = None,
    drive_id: str | None = None,
    group_ids: list[str] | None = None,
    after: str | None = None,
    before: str | None = None,
) -> list[dict]:
    """Return SharePoint search results as plain dicts.

    `after` and `before` accept ISO 8601 date strings.
    `group_ids` scopes the search to specific M365 group drives.
    Raises on API error. Callers are responsible for handling exceptions.
    """
    search_query = query

    filters = []
    if before:
        filters.append(f"created<={before}")
    if after:
        filters.append(f"created>={after}")
    if filters:
        search_query += " " + " ".join(filters)

    if group_ids:
        group_filter = " OR ".join([f"GroupId:{gid}" for gid in group_ids])
        search_query = f"({search_query}) AND ({group_filter})"

    search_options = graph_search.SearchOptions(
        query_string=search_query,
        sort_by=sort_by,
        descending=descending,
        page_size=page_size,
        region=region,
        drive_id=drive_id,
    )

    items: list[dict] = []
    async for _item in graph_search.search_entities(
        context.graph_client,
        entity_types=[EntityType.DriveItem],
        options=search_options,
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
        created = (
            drive_item.created_date_time.strftime("%Y-%m-%d")
            if drive_item.created_date_time
            else ""
        )

        items.append({
            "drive_id": (
                drive_item.parent_reference.drive_id if drive_item.parent_reference else None
            ),
            "item_id": drive_item.id,
            "name": drive_item.name,
            "size": drive_item.size,
            "size_str": size_str,
            "web_url": drive_item.web_url,
            "e_tag": drive_item.e_tag,
            "c_tag": drive_item.c_tag,
            "description": drive_item.description,
            "created_datetime": (
                drive_item.created_date_time.isoformat() if drive_item.created_date_time else None
            ),
            "last_modified_datetime": (
                drive_item.last_modified_date_time.isoformat()
                if drive_item.last_modified_date_time
                else None
            ),
            "created": created,
            "last_modified": (
                drive_item.last_modified_date_time.strftime("%Y-%m-%d")
                if drive_item.last_modified_date_time
                else None
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
            "last_modified_by_email": (
                drive_item.last_modified_by.user.additional_data.get("email")
                if drive_item.last_modified_by and drive_item.last_modified_by.user
                else None
            ),
            "last_modified_by_id": (
                drive_item.last_modified_by.user.id
                if drive_item.last_modified_by and drive_item.last_modified_by.user
                else None
            ),
            "parent_id": (
                drive_item.parent_reference.id if drive_item.parent_reference else None
            ),
            "parent_name": (
                drive_item.parent_reference.name if drive_item.parent_reference else None
            ),
            "parent_path": (
                drive_item.parent_reference.path if drive_item.parent_reference else None
            ),
            "site_id": (
                drive_item.parent_reference.site_id if drive_item.parent_reference else None
            ),
            "drive_type": (
                drive_item.parent_reference.drive_type if drive_item.parent_reference else None
            ),
            "mime_type": (drive_item.file.mime_type if drive_item.file else None),
            "is_folder": drive_item.folder is not None,
            "child_count": (
                drive_item.folder.child_count if drive_item.folder else None
            ),
            "download_url": (
                drive_item.additional_data.get("@microsoft.graph.downloadUrl")
                if drive_item.additional_data
                else None
            ),
        })

    return items


def add_arguments(parser: "argparse.ArgumentParser"):
    parser.set_defaults(uses_time_bounds=True)
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
        "--visibility",
        type=str,
        choices=["Private", "Public"],
        default=None,
        help="Filter groups by visibility (use with --my-groups).",
    )


@handle_graph_errors
async def run_with_arguments(
    context: "GraphContext", args: "argparse.Namespace"
) -> int:

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
    drive_id = getattr(args, "drive_id", None)

    group_ids: list[str] | None = None
    if args.my_groups:
        groups = await get_user_m365_groups(
            context.graph_client, visibility=args.visibility, teams_only=True
        )
        if not groups:
            logger.warning("No Microsoft 365 groups found, no results will be returned")
            return 0
        group_ids = [g.id for g in groups]
        logger.info(f"Scoping search to {len(group_ids)} user groups")

    if drive_id:
        logger.info(f"Scoping search to Drive ID: {drive_id}")

    logger.info(f"Search query: {search_query}")

    save_dir = None
    if args.save:
        save_dir = Path(args.save)
        save_dir.mkdir(parents=True, exist_ok=True)

    cached_items: list[dict] = []
    downloaded = 0
    failed = 0

    if not save_dir and not context.json_output:
        console.print("[bold]Search results[/bold]")
        console.rule()

    try:
        cached_items = await fetch(
            context,
            query=search_query,
            region=region,
            drive_id=drive_id,
            group_ids=group_ids,
            after=after_iso,
            before=before_iso,
        )

        for count, item in enumerate(cached_items, 1):
            if not save_dir and not context.json_output and not context.ndjson_output:
                console.print(
                    f"  [dim]{count:>4}.[/dim]  {item['name']}  "
                    f"[dim]{item['author']}[/dim]  [cyan]{item['size_str']}[/cyan]  [dim]{item['created']}[/dim]"
                )

            if context.ndjson_output:
                output.print_ndjson_item(item)

            if save_dir and item["drive_id"] and item["item_id"]:
                try:
                    safe_filename = (item["name"] or "unknown").replace("/", "_").replace("..", "_")
                    file_path = save_dir / safe_filename
                    info_path = save_dir / f"{safe_filename}_info.json"

                    if file_path.exists() and info_path.exists():
                        logger.debug(f" Skipping (already exists): {file_path.name}")
                        downloaded += 1
                        continue

                    file_stream = (
                        await context.graph_client.drives.by_drive_id(item["drive_id"])
                        .items.by_drive_item_id(item["item_id"])
                        .content.get()
                    )

                    if file_stream:
                        with open(file_path, "wb") as f:
                            f.write(file_stream)

                        metadata = {
                            "id": item["item_id"],
                            "name": item["name"],
                            "size": item["size"],
                            "created_datetime": item["created_datetime"],
                            "last_modified_datetime": item["last_modified_datetime"],
                            "web_url": item["web_url"],
                            "created_by": {
                                "user": {
                                    "display_name": item["author"],
                                    "email": item["created_by_email"],
                                    "id": item["created_by_id"],
                                }
                            } if item["author"] else None,
                            "last_modified_by": {
                                "user": {
                                    "display_name": item["last_modified_by"],
                                    "email": item["last_modified_by_email"],
                                    "id": item["last_modified_by_id"],
                                }
                            } if item["last_modified_by"] else None,
                            "parent_reference": {
                                "drive_id": item["drive_id"],
                                "drive_type": item["drive_type"],
                                "id": item["parent_id"],
                                "name": item["parent_name"],
                                "path": item["parent_path"],
                                "site_id": item["site_id"],
                            },
                            "file": {"mime_type": item["mime_type"]} if item["mime_type"] else None,
                        }

                        with open(info_path, "w", encoding="utf-8") as f:
                            json.dump(metadata, f, indent=2, ensure_ascii=False)

                        logger.debug(f"Saved: {file_path.name}")
                        downloaded += 1
                    else:
                        logger.warning(f"Empty file: {item['name']}")

                except Exception as e:
                    logger.error(f"Failed to download {item['name']}: {e}")
                    failed += 1

    except KeyboardInterrupt:
        if cached_items:
            logger.info(f"Interrupted. {len(cached_items)} result(s) cached.")
    finally:
        if cached_items:
            cache.save_results(cached_items, key="sharepoint")

    if not cached_items:
        logger.info("No results found.")
        if context.json_output:
            output.print_json([])
        return 0

    if save_dir:
        logger.info(f"{downloaded} items saved to: {save_dir.absolute()}")
        if failed > 0:
            logger.warning(f"Failed: {failed}")

    if context.json_output:
        output.print_json(cached_items)
    # ndjson items streamed inline

    return 0
