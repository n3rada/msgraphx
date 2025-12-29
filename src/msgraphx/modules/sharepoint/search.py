# msgraphx/modules/sharepoint/search.py

# Built-in imports
import json
from pathlib import Path
from typing import TYPE_CHECKING

# External library imports
from loguru import logger
from msgraph import GraphServiceClient
from msgraph.generated.models.entity_type import EntityType

# Local library imports
from ...core import graph_search
from ...utils.errors import handle_graph_errors
from ...utils.dates import parse_date_string


if TYPE_CHECKING:
    import argparse
    from msgraphx.core.context import GraphContext

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


@handle_graph_errors
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


@handle_graph_errors
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

    # Create save directory if specified
    save_dir = None
    if args.save:
        save_dir = Path(args.save)
        save_dir.mkdir(parents=True, exist_ok=True)

    count = 0
    downloaded = 0
    failed = 0

    async for drive_item in graph_search.search_entities(
        context.graph_client,
        entity_types=[EntityType.DriveItem],
        options=search_options,
    ):

        logger.trace(drive_item.__dict__)

        # Process each DriveItem
        logger.info(
            f"📄 {drive_item.name} (Created by: {drive_item.created_by.user.display_name})"
        )
        count += 1

        # Download file if --save is specified
        if save_dir:
            try:
                # Sanitize filename to avoid path traversal
                safe_filename = drive_item.name.replace("/", "_").replace("..", "_")
                file_path = save_dir / safe_filename
                info_path = save_dir / f"{safe_filename}_info.json"

                # Check if file already exists - skip download if it does
                if file_path.exists() and info_path.exists():
                    logger.debug(f"⏭️  Skipping (already exists): {file_path.name}")
                    downloaded += 1  # Count as downloaded
                    continue

                # Get file content stream only if file doesn't exist
                file_stream = (
                    await context.graph_client.drives.by_drive_id(
                        drive_item.parent_reference.drive_id
                    )
                    .items.by_drive_item_id(drive_item.id)
                    .content.get()
                )

                if file_stream:
                    # Write file
                    with open(file_path, "wb") as f:
                        f.write(file_stream)

                    # Save metadata as JSON
                    metadata = {
                        "id": drive_item.id,
                        "name": drive_item.name,
                        "size": drive_item.size,
                        "created_datetime": (
                            drive_item.created_date_time.isoformat()
                            if drive_item.created_date_time
                            else None
                        ),
                        "last_modified_datetime": (
                            drive_item.last_modified_date_time.isoformat()
                            if drive_item.last_modified_date_time
                            else None
                        ),
                        "web_url": drive_item.web_url,
                        "created_by": (
                            {
                                "user": {
                                    "display_name": (
                                        drive_item.created_by.user.display_name
                                        if drive_item.created_by
                                        and drive_item.created_by.user
                                        else None
                                    ),
                                    "email": (
                                        drive_item.created_by.user.email
                                        if drive_item.created_by
                                        and drive_item.created_by.user
                                        else None
                                    ),
                                    "id": (
                                        drive_item.created_by.user.id
                                        if drive_item.created_by
                                        and drive_item.created_by.user
                                        else None
                                    ),
                                }
                            }
                            if drive_item.created_by
                            else None
                        ),
                        "last_modified_by": (
                            {
                                "user": {
                                    "display_name": (
                                        drive_item.last_modified_by.user.display_name
                                        if drive_item.last_modified_by
                                        and drive_item.last_modified_by.user
                                        else None
                                    ),
                                    "email": (
                                        drive_item.last_modified_by.user.email
                                        if drive_item.last_modified_by
                                        and drive_item.last_modified_by.user
                                        else None
                                    ),
                                    "id": (
                                        drive_item.last_modified_by.user.id
                                        if drive_item.last_modified_by
                                        and drive_item.last_modified_by.user
                                        else None
                                    ),
                                }
                            }
                            if drive_item.last_modified_by
                            else None
                        ),
                        "parent_reference": (
                            {
                                "drive_id": (
                                    drive_item.parent_reference.drive_id
                                    if drive_item.parent_reference
                                    else None
                                ),
                                "drive_type": (
                                    drive_item.parent_reference.drive_type
                                    if drive_item.parent_reference
                                    else None
                                ),
                                "id": (
                                    drive_item.parent_reference.id
                                    if drive_item.parent_reference
                                    else None
                                ),
                                "name": (
                                    drive_item.parent_reference.name
                                    if drive_item.parent_reference
                                    else None
                                ),
                                "path": (
                                    drive_item.parent_reference.path
                                    if drive_item.parent_reference
                                    else None
                                ),
                                "site_id": (
                                    drive_item.parent_reference.site_id
                                    if drive_item.parent_reference
                                    else None
                                ),
                            }
                            if drive_item.parent_reference
                            else None
                        ),
                        "file": (
                            {
                                "mime_type": (
                                    drive_item.file.mime_type
                                    if drive_item.file
                                    else None
                                ),
                            }
                            if drive_item.file
                            else None
                        ),
                    }

                    with open(info_path, "w", encoding="utf-8") as f:
                        json.dump(metadata, f, indent=2, ensure_ascii=False)

                    logger.debug(f"✅ Saved: {file_path.name}")
                    downloaded += 1
                else:
                    logger.warning(f"⚠️ Empty file: {drive_item.name}")

            except Exception as e:
                logger.error(f"❌ Failed to download {drive_item.name}: {e}")
                failed += 1

    if count > 0:

        logger.info(f"📊 Total files found: {count}")
        if save_dir:
            logger.info(
                f"💾 {downloaded} items have been saved to: {save_dir.absolute()}"
            )
            if failed > 0:
                logger.warning(f"⚠️ Failed: {failed}")

    return 0
