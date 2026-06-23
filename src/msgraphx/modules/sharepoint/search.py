# msgraphx/modules/sharepoint/search.py

# Built-in imports
from __future__ import annotations

import argparse
import json
from pathlib import Path

# External library imports
from loguru import logger
from msgraph.generated.models.search_content import SearchContent
from msgraph.generated.models.share_point_one_drive_options import SharePointOneDriveOptions

# Local library imports
from .groups import get_user_m365_groups
from ...core.drive_search import DriveSearchBase
from ...core.context import GraphContext
from ...utils import cache, output
from ...utils.console import console
from ...utils.errors import handle_graph_errors


async def _resolve_site(context: GraphContext, site_ref: str) -> str | None:
    """Resolve a site name or Graph site ID to its SharePoint web URL.

    Accepts:
      - A SharePoint URL (https://tenant.sharepoint.com/sites/Name) — used as-is
      - A site name/slug — resolved via GET /sites?$search=<name>
      - A full Graph site ID (contains commas) — resolved via GET /sites/{id}
    """
    from msgraph.generated.sites.sites_request_builder import SitesRequestBuilder
    from kiota_abstractions.base_request_configuration import RequestConfiguration
    SitesRequestBuilderGetQueryParameters = SitesRequestBuilder.SitesRequestBuilderGetQueryParameters

    # Full SharePoint URL — use directly as KQL Path:
    if site_ref.startswith("http"):
        return site_ref

    # Graph site ID (hostname,siteId,webId) — resolve to web URL
    if site_ref.count(",") >= 2:
        try:
            site = await context.graph_client.sites.by_site_id(site_ref).get()
            return site.web_url if site else None
        except Exception as exc:
            logger.error(f"Failed to resolve site ID {site_ref}: {exc}")
            return None

    # Name/slug: search via GET /sites?$search=<name>
    try:
        config = RequestConfiguration(
            query_parameters=SitesRequestBuilderGetQueryParameters(search=site_ref, top=5),
        )
        resp = await context.graph_client.sites.get(request_configuration=config)
        sites = resp.value if resp and resp.value else []
        if not sites:
            logger.error(f"No site found matching '{site_ref}'.")
            return None
        if len(sites) > 1:
            matches = ", ".join(f"{s.display_name} ({s.web_url})" for s in sites[:5])
            logger.warning(f"Multiple sites match '{site_ref}': {matches}")
            logger.warning("Using first match. Pass a full URL or site ID to be precise.")
        return sites[0].web_url
    except Exception as exc:
        logger.error(f"Failed to search for site '{site_ref}': {exc}")
        return None

HUNT_QUERIES: dict[str, str] = {
    "scripts": (
        "((filetype:ps1 OR filetype:sh OR filetype:bat OR filetype:cmd OR "
        "filetype:py OR filetype:rb OR filetype:pl OR filetype:ts) "
        'OR ("#!" OR "#!/" OR "param(" OR "function " OR "Import-Module"))'
    ),
    "credentials": (
        "((filetype:key OR filetype:pem OR filetype:crt OR filetype:cer OR filetype:kdbx "
        "OR filetype:pfx) "
        "OR ((filetype:env OR filetype:cfg OR filetype:yaml OR filetype:yml OR filetype:secret) "
        'AND ("password" OR "passwd" OR "secret" OR "api_key" OR "access_key" OR "client_secret")))'
    ),
    "ssh": (
        "(filetype:pub OR filetype:pem OR filename:id_rsa OR filename:id_ecdsa OR "
        "filename:id_ed25519 OR filename:id_dsa OR filename:authorized_keys OR "
        'filename:known_hosts OR filename:.ssh OR "BEGIN RSA PRIVATE KEY" OR '
        '"BEGIN OPENSSH PRIVATE KEY" OR "BEGIN DSA PRIVATE KEY" OR '
        '"BEGIN EC PRIVATE KEY" OR "BEGIN PRIVATE KEY" OR "ssh-rsa" OR '
        '"ssh-ed25519" OR "ecdsa-sha2")'
    ),
    "mssql": (
        "( (filetype:mdf OR filetype:ldf OR filetype:udl OR filetype:sql) "
        'OR ("Data Source=" AND "Initial Catalog=") '
        'OR ("connection string" AND ("User ID=" OR "Uid=" OR "Password=")) '
        'OR (filename:web.config AND "connectionStrings") '
        'OR (filename:appsettings.json AND "ConnectionStrings") '
        'OR (filetype:sql AND ("CREATE TABLE" OR "INSERT INTO" OR "ALTER TABLE")) )'
    ),
    "office": (
        "filetype:doc OR filetype:docx OR filetype:rtf OR filetype:odt OR "
        "filetype:xls OR filetype:xlsx OR filetype:csv OR filetype:ods OR "
        "filetype:ppt OR filetype:pptx OR filetype:pdf OR filetype:msg OR filetype:eml"
    ),
    "backups": (
        "(filetype:bak OR filetype:vhd OR filetype:vmdk OR filetype:ova) "
        "OR ((filetype:zip OR filetype:rar OR filetype:7z OR filetype:gz OR filetype:tar) "
        'AND ("backup" OR "dump" OR "dbbackup" OR "backup_")) '
        'OR (filetype:sql AND "backup")'
    ),
    "configs": (
        "((filetype:conf OR filetype:ini OR filetype:psd1 OR filetype:reg) "
        "OR filename:appsettings.json OR filename:web.config OR filename:settings.json) "
        'AND ("connection" OR "password" OR "secret" OR "token" OR "credentials")'
    ),
    "infra": (
        "(filetype:dockerfile OR filename:Dockerfile OR filetype:compose OR "
        '(filetype:tf OR filename:terraform) AND ("resource" OR "provider" OR "module"))'
    ),
    "network": "filetype:pcap OR filetype:cap OR filetype:har",
}


class SharePointSearch(DriveSearchBase):
    """Search SharePoint team sites (SharedContent for app-only tokens)."""

    CACHE_KEY = "sharepoint"
    LABEL = "SharePoint"
    HUNT_QUERIES = HUNT_QUERIES
    # For app-only: restrict to SharePoint team sites only.
    # For delegated: this is ignored; the user sees everything they have access to.
    SCOPE = SharePointOneDriveOptions(include_content=SearchContent.SharedContent)


# Preserve module-level public API expected by the module contract.
async def fetch(
    context: GraphContext,
    query: str = "*",
    region: str | None = None,
    drive_id: str | None = None,
    group_ids: list[str] | None = None,
    after: str | None = None,
    before: str | None = None,
) -> list[dict]:
    return await SharePointSearch.fetch(
        context,
        query=query,
        after=after,
        before=before,
        drive_id=drive_id,
        group_ids=group_ids,
    )


def add_arguments(parser: argparse.ArgumentParser) -> None:
    SharePointSearch.add_arguments(parser)

    parser.add_argument(
        "--site",
        type=str,
        default=None,
        metavar="SITE",
        help=(
            "Restrict search to a specific SharePoint site. "
            "Accepts a site name/slug (searched via $search), a full SharePoint URL, "
            "or a Graph site ID."
        ),
    )
    parser.add_argument(
        "--my-groups",
        action="store_true",
        help="Search only in the current user's Microsoft 365 groups.",
    )
    parser.add_argument(
        "--visibility",
        type=str,
        choices=["Private", "Public"],
        default=None,
        help="Filter groups by visibility (use with --my-groups).",
    )
    parser.add_argument(
        "--scope",
        choices=["sharepoint", "onedrive"],
        default=None,
        help=(
            "Restrict search to SharePoint team sites or personal OneDrive drives. "
            "Requires app-only (client credentials) auth — the API ignores this for delegated tokens."
        ),
    )


@handle_graph_errors
async def run_with_arguments(
    context: GraphContext, args: argparse.Namespace
) -> int:
    scope = getattr(args, "scope", None)
    if scope and not context.is_app_only:
        logger.error(
            "--scope requires application permissions (app-only token). "
            "The Graph API silently ignores SharePointOneDriveOptions for delegated tokens."
        )
        return 1

    _SCOPE_MAP = {
        "sharepoint": SharePointOneDriveOptions(include_content=SearchContent.SharedContent),
        "onedrive": SharePointOneDriveOptions(include_content=SearchContent.PrivateContent),
    }
    resolved_scope = _SCOPE_MAP.get(scope) if scope else None
    resolved_label = {"sharepoint": "SharePoint", "onedrive": "OneDrive"}.get(scope, "SharePoint / OneDrive")

    group_ids: list[str] | None = None
    if getattr(args, "my_groups", False):
        # Fetch ALL M365 Unified groups, not just Teams-provisioned ones.
        # Every Unified group owns a SharePoint site regardless of whether
        # a Teams workspace was ever added on top of it.
        groups = await get_user_m365_groups(
            context.graph_client,
            visibility=getattr(args, "visibility", None),
            teams_only=False,
        )
        if not groups:
            logger.warning("No M365 Unified groups found.")
            return 0
        group_ids = [g.id for g in groups]
        logger.info(f"Scoping search to {len(group_ids)} M365 Unified groups")

    args._group_ids = group_ids
    args._resolved_scope = resolved_scope
    args._resolved_label = resolved_label

    site_path: str | None = None
    site_ref = getattr(args, "site", None)
    if site_ref:
        site_path = await _resolve_site(context, site_ref)
        if not site_path:
            return 1
        logger.info(f"Scoping search to site: {site_path}")
    args._site_path = site_path

    save_dir = None
    if args.save:
        save_dir = Path(args.save)
        save_dir.mkdir(parents=True, exist_ok=True)

    if save_dir:
        return await _run_with_save(context, args, save_dir)

    return await SharePointSearch.run_with_arguments(context, args, scope=resolved_scope, label=resolved_label)


async def _run_with_save(
    context: GraphContext, args: argparse.Namespace, save_dir: Path
) -> int:
    """Run search and download every result directly to save_dir."""
    from ...utils.dates import parse_date_string

    search_query = args.query
    hunt = getattr(args, "hunt", None)
    if args.filetype:
        search_query = f"filetype:{args.filetype} {search_query}".strip()
    elif hunt and hunt in HUNT_QUERIES:
        search_query = HUNT_QUERIES[hunt]
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

    items = await SharePointSearch.fetch(
        context,
        query=search_query,
        after=after_iso,
        before=before_iso,
        drive_id=getattr(args, "drive_id", None),
        group_ids=getattr(args, "_group_ids", None),
    )

    if not items:
        logger.info("No results found.")
        return 0

    cache.save_results(items, key="sharepoint", identity=context.identity_hash)

    downloaded = 0
    failed = 0

    for item in items:
        if not item["drive_id"] or not item["item_id"]:
            continue
        try:
            safe_name = (item["name"] or "unknown").replace("/", "_").replace("..", "_")
            file_path = save_dir / safe_name
            info_path = save_dir / f"{safe_name}_info.json"

            if file_path.exists() and info_path.exists():
                logger.debug(f"Skipping (exists): {safe_name}")
                downloaded += 1
                continue

            file_stream = (
                await context.graph_client.drives.by_drive_id(item["drive_id"])
                .items.by_drive_item_id(item["item_id"])
                .content.get()
            )
            if file_stream:
                file_path.write_bytes(file_stream)
                info_path.write_text(
                    json.dumps(item, indent=2, ensure_ascii=False), encoding="utf-8"
                )
                downloaded += 1
            else:
                logger.warning(f"Empty file: {item['name']}")
        except Exception as exc:
            logger.error(f"Failed to download {item['name']}: {exc}")
            failed += 1

    logger.info(f"{downloaded} item(s) saved to {save_dir.absolute()}")
    if failed:
        logger.warning(f"Failed: {failed}")

    if context.json_output:
        output.print_json(items)

    return 0
