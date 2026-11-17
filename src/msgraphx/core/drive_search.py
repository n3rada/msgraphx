# msgraphx/core/drive_search.py
#
# Base class for DriveItem search modules (SharePoint, OneDrive).
# Subclasses declare CACHE_KEY, SCOPE, and HUNT_QUERIES; the base
# handles argument wiring, result shaping, caching, and output.

from __future__ import annotations

import argparse
from collections.abc import Sequence
from typing import Any, ClassVar

from loguru import logger
from msgraph.generated.models.drive_item import DriveItem
from msgraph.generated.models.entity_type import EntityType
from msgraph.generated.models.share_point_one_drive_options import SharePointOneDriveOptions

from . import graph_search
from ..core.context import GraphContext
from ..utils import cache, output
from ..utils.console import console
from ..utils.dates import parse_date_string
from ..utils.errors import handle_graph_errors


def drive_item_to_dict(item: DriveItem) -> dict[str, Any]:
    """Flatten a DriveItem into a JSON-serializable record.

    Args:
        item: A drive item returned by the SDK, typically a search hit.

    Returns:
        A dict with the item's IDs, name, sizes (raw and human-readable
        `size_str`), URLs, ISO and `YYYY-MM-DD` timestamps, creator and
        last modifier, parent reference fields, MIME type, folder details,
        and pre-authenticated `download_url` when present. Missing values
        are `None`, except `author` (`?`) and `created` (empty string).
    """
    author = (
        item.created_by.user.display_name
        if item.created_by and item.created_by.user
        else "?"
    )
    size_bytes = item.size or 0
    if size_bytes >= 1_048_576:
        size_str = f"{size_bytes / 1_048_576:.1f} MB"
    elif size_bytes >= 1024:
        size_str = f"{size_bytes / 1024:.1f} KB"
    else:
        size_str = f"{size_bytes} B"

    return {
        "drive_id": (
            item.parent_reference.drive_id if item.parent_reference else None
        ),
        "item_id": item.id,
        "name": item.name,
        "size": item.size,
        "size_str": size_str,
        "web_url": item.web_url,
        "e_tag": item.e_tag,
        "c_tag": item.c_tag,
        "description": item.description,
        "created_datetime": (
            item.created_date_time.isoformat() if item.created_date_time else None
        ),
        "last_modified_datetime": (
            item.last_modified_date_time.isoformat()
            if item.last_modified_date_time
            else None
        ),
        "created": (
            item.created_date_time.strftime("%Y-%m-%d")
            if item.created_date_time
            else ""
        ),
        "last_modified": (
            item.last_modified_date_time.strftime("%Y-%m-%d")
            if item.last_modified_date_time
            else None
        ),
        "author": author,
        "created_by_email": (
            item.created_by.user.additional_data.get("email")
            if item.created_by and item.created_by.user
            else None
        ),
        "created_by_id": (
            item.created_by.user.id
            if item.created_by and item.created_by.user
            else None
        ),
        "last_modified_by": (
            item.last_modified_by.user.display_name
            if item.last_modified_by and item.last_modified_by.user
            else None
        ),
        "last_modified_by_email": (
            item.last_modified_by.user.additional_data.get("email")
            if item.last_modified_by and item.last_modified_by.user
            else None
        ),
        "last_modified_by_id": (
            item.last_modified_by.user.id
            if item.last_modified_by and item.last_modified_by.user
            else None
        ),
        "parent_id": (
            item.parent_reference.id if item.parent_reference else None
        ),
        "parent_name": (
            item.parent_reference.name if item.parent_reference else None
        ),
        "parent_path": (
            item.parent_reference.path if item.parent_reference else None
        ),
        "site_id": (
            item.parent_reference.site_id if item.parent_reference else None
        ),
        "drive_type": (
            item.parent_reference.drive_type if item.parent_reference else None
        ),
        "mime_type": (item.file.mime_type if item.file else None),
        "is_folder": item.folder is not None,
        "child_count": (item.folder.child_count if item.folder else None),
        "download_url": (
            item.additional_data.get("@microsoft.graph.downloadUrl")
            if item.additional_data
            else None
        ),
    }


class DriveSearchBase:
    """A DriveItem search command, specialized by SharePoint and OneDrive.

    Subclasses override the class attributes; the base class provides
    argument wiring, query building, result shaping, caching, and output.

    Attributes:
        CACHE_KEY: Key for `utils.cache`, e.g. `sharepoint` or `onedrive`.
        SCOPE: Search scope applied for app-only tokens; `None` searches all
            drives.
        HUNT_QUERIES: Predefined KQL queries offered by `--hunt`, by name.
        LABEL: Human-readable name used in console output.
    """

    CACHE_KEY: ClassVar[str] = "drives"
    SCOPE: ClassVar[SharePointOneDriveOptions | None] = None
    HUNT_QUERIES: ClassVar[dict[str, str]] = {}
    LABEL: ClassVar[str] = "Drive"

    # ------------------------------------------------------------------
    # Core fetch
    # ------------------------------------------------------------------

    @classmethod
    async def fetch(
        cls,
        context: GraphContext,
        query: str = "*",
        after: str | None = None,
        before: str | None = None,
        drive_id: str | None = None,
        group_ids: Sequence[str] | None = None,
        scope: SharePointOneDriveOptions | None = None,
        site_path: str | None = None,
    ) -> list[dict[str, Any]]:
        """Search drive items and return them as flat records.

        Args:
            context: The shared Graph context.
            query: KQL query string.
            after: Keep items created at or after this ISO 8601 timestamp.
            before: Keep items created at or before this ISO 8601 timestamp.
            drive_id: Restrict results to this drive.
            group_ids: Restrict results to these Microsoft 365 group IDs.
            scope: Search scope for app-only tokens, overriding `SCOPE`.
                Ignored for delegated tokens.
            site_path: Restrict results to this site URL.

        Returns:
            One `drive_item_to_dict` record per matching drive item.

        Raises:
            ForbiddenGraphError: If the search API returns 403 Forbidden.
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
            group_filter = " OR ".join(f"GroupId:{gid}" for gid in group_ids)
            search_query = f"({search_query}) AND ({group_filter})"

        # share_point_one_drive_options only applies to app-only tokens.
        # Use caller-supplied scope override first, then fall back to class-level SCOPE.
        resolved = (scope if scope is not None else cls.SCOPE) if context.is_app_only else None

        options = graph_search.SearchOptions(
            query_string=search_query,
            page_size=500,
            region=context.region if context.is_app_only else None,
            drive_id=drive_id,
            share_point_one_drive_options=resolved,
            site_path=site_path,
        )

        items: list[dict] = []
        async for raw in graph_search.search_entities(
            context.graph_client,
            entity_types=[EntityType.DriveItem],
            options=options,
        ):
            if isinstance(raw, DriveItem):
                items.append(drive_item_to_dict(raw))

        return items

    # ------------------------------------------------------------------
    # CLI argument wiring
    # ------------------------------------------------------------------

    @classmethod
    def add_arguments(cls, parser: argparse.ArgumentParser) -> None:
        """Register the search command's arguments.

        Adds the positional KQL `query`, `--filetype`, and `--hunt` when
        `HUNT_QUERIES` is non-empty, and marks the command as honoring the
        global `--before`/`--after` bounds.

        Args:
            parser: The subcommand parser to configure.
        """
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
            metavar="EXT",
            help="Filter by file extension (e.g. pdf, docx, xlsx).",
        )
        if cls.HUNT_QUERIES:
            parser.add_argument(
                "--hunt",
                choices=cls.HUNT_QUERIES.keys(),
                help="Predefined hunt query.",
            )

    # ------------------------------------------------------------------
    # CLI run
    # ------------------------------------------------------------------

    @classmethod
    @handle_graph_errors
    async def run_with_arguments(
        cls,
        context: GraphContext,
        args: argparse.Namespace,
        scope: SharePointOneDriveOptions | None = None,
        label: str | None = None,
    ) -> int:
        """Run the search from parsed arguments and print or cache the results.

        `--filetype` takes precedence over `--hunt`. Results are cached under
        `CACHE_KEY` for later download by index, then written as JSON, as
        NDJSON, or as a console table depending on the output mode.

        Args:
            context: The shared Graph context.
            args: Parsed arguments: `query`, `filetype`, `hunt`, `before`,
                `after`, and optionally `drive_id`, `_group_ids`, and
                `_site_path` set by calling modules.
            scope: Search scope for app-only tokens, overriding `SCOPE`.
            label: Console heading, overriding `LABEL`.

        Returns:
            0 on success, including no results; 1 if a date bound is invalid.

        Raises:
            AuthenticationError: If the token is invalid or expired.
            ForbiddenGraphError: If the search API returns 403 Forbidden.
        """
        search_query = args.query

        hunt = getattr(args, "hunt", None)
        if args.filetype:
            search_query = f"filetype:{args.filetype} {search_query}".strip()
        elif hunt and hunt in cls.HUNT_QUERIES:
            search_query = cls.HUNT_QUERIES[hunt]
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

        drive_id = getattr(args, "drive_id", None)
        group_ids: list[str] | None = getattr(args, "_group_ids", None)
        site_path: str | None = getattr(args, "_site_path", None)

        logger.info(f"Search query: {search_query}")

        items = await cls.fetch(
            context,
            query=search_query,
            after=after_iso,
            before=before_iso,
            drive_id=drive_id,
            group_ids=group_ids,
            scope=scope,
            site_path=site_path,
        )

        if not items:
            logger.info("No results found.")
            if context.json_output:
                output.print_json([])
            return 0

        cache.save_results(items, key=cls.CACHE_KEY, identity=context.identity_hash)

        if context.json_output:
            output.print_json(items)
            return 0

        if context.ndjson_output:
            for item in items:
                output.print_ndjson_item(item)
            return 0

        console.print(f"[bold]{label or cls.LABEL} search results[/bold] ({len(items)})")
        console.rule()
        for count, item in enumerate(items, 1):
            console.print(
                f"  [dim]{count:>4}.[/dim]  {item['name']}  "
                f"[dim]{item['author']}[/dim]  [cyan]{item['size_str']}[/cyan]  "
                f"[dim]{item['created']}[/dim]"
            )

        return 0
