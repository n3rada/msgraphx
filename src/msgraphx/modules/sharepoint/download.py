# msgraphx/modules/sharepoint/download.py

# Built-in imports
from __future__ import annotations

import argparse
import asyncio
from collections.abc import Sequence
from pathlib import Path

# External library imports
from loguru import logger
from msgraph.generated.models.drive_item import DriveItem
from msgraph.graph_service_client import GraphServiceClient

# Local library imports
from ...core.context import GraphContext
from ...utils import cache, pagination
from ...utils.errors import handle_graph_errors, is_graph_auth_error
from ...utils.roles import require_scopes


@handle_graph_errors
async def download_drive_item(
    graph_client: GraphServiceClient,
    drive_id: str,
    item: DriveItem,
    output_dir: Path,
    base_path: str = "",
    semaphore: asyncio.Semaphore | None = None,
    skip_existing: bool = True,
) -> int:
    """Download a drive item, recursing into folders.

    Folder children are downloaded concurrently. Per-item failures are logged
    and do not stop the rest of the download, except authentication failures,
    which abort it. Items without a name or ID are skipped with a warning.

    Args:
        graph_client: Authenticated Graph client.
        drive_id: ID of the drive that holds the item.
        item: The file or folder to download.
        output_dir: Local directory that mirrors the drive root.
        base_path: Path of the item's parent relative to the drive root; ""
            for root-level items.
        semaphore: Limits concurrent file downloads. None means no limit.
        skip_existing: Skip local files whose size already matches the
            remote size (resume mode).

    Returns:
        The number of files downloaded or skipped as already present.

    Raises:
        AuthenticationError: The Graph token is invalid or expired.
    """
    downloaded_count = 0

    if not item.name or not item.id:
        logger.warning(
            f"Skipping drive item without name or ID under '{base_path or '/'}' "
            f"(id={item.id!r}, name={item.name!r})"
        )
        return 0

    # Build the full path
    current_path = f"{base_path}/{item.name}" if base_path else item.name

    if item.folder:
        # It's a folder - recurse into it
        logger.debug(f"Entering folder: {current_path}")
        try:
            # Get all children of this folder (with pagination)
            children: list[DriveItem] = await pagination.collect_all(
                graph_client.drives.by_drive_id(drive_id)
                .items.by_drive_item_id(item.id)
                .children
            )

            if children:
                # Download all children concurrently
                tasks = [
                    download_drive_item(
                        graph_client,
                        drive_id,
                        child,
                        output_dir,
                        current_path,
                        semaphore,
                        skip_existing,
                    )
                    for child in children
                ]
                results = await asyncio.gather(*tasks, return_exceptions=True)
                downloaded_count += _sum_results(results, f"folder {current_path}")

        except Exception as e:
            if _is_auth_failure(e):
                raise
            logger.error(f"Failed to list folder {current_path}: {e}")

    elif item.file:
        # It's a file - download it with semaphore to limit concurrency
        if semaphore:
            async with semaphore:
                downloaded_count = await _download_file(
                    graph_client,
                    drive_id,
                    item,
                    output_dir,
                    current_path,
                    skip_existing,
                )
        else:
            downloaded_count = await _download_file(
                graph_client, drive_id, item, output_dir, current_path, skip_existing
            )

    return downloaded_count


def _is_auth_failure(exc: BaseException) -> bool:
    """Return True if `exc` means the Graph token is invalid or expired.

    Accepts the `BaseException` values `asyncio.gather` returns, such as
    `CancelledError`, which are never auth failures.

    Args:
        exc: The exception to inspect.
    """
    return isinstance(exc, Exception) and is_graph_auth_error(exc)


def _sum_results(results: Sequence[int | BaseException], where: str) -> int:
    """Total the per-item counts from `asyncio.gather(return_exceptions=True)`.

    Args:
        results: Gathered results: a file count or the exception a task raised.
        where: Location used in error log messages, e.g. "folder Docs".

    Returns:
        The sum of the integer results. Other exceptions are logged.

    Raises:
        AuthenticationError: A task failed because the token is invalid or
            expired; the first such error is re-raised to stop the dump.
    """
    total = 0
    auth_error: BaseException | None = None
    for result in results:
        if not isinstance(result, BaseException):
            total += result
        elif _is_auth_failure(result):
            auth_error = auth_error or result
        else:
            logger.error(f"Error in {where}: {result}")
    if auth_error is not None:
        raise auth_error
    return total


@handle_graph_errors
async def _download_file(
    graph_client: GraphServiceClient,
    drive_id: str,
    item: DriveItem,
    output_dir: Path,
    current_path: str,
    skip_existing: bool = True,
) -> int:
    """Download one file to `output_dir / current_path`.

    Errors are logged rather than raised, except authentication failures.

    Args:
        graph_client: Authenticated Graph client.
        drive_id: ID of the drive that holds the file.
        item: The file to download.
        output_dir: Local directory that mirrors the drive root.
        current_path: File path relative to the drive root.
        skip_existing: Skip the download if a local file of the same size
            already exists.

    Returns:
        1 if the file was downloaded or skipped as already present, 0 if it
        had no content or the download failed.

    Raises:
        AuthenticationError: The Graph token is invalid or expired.
    """
    if not item.id:
        logger.warning(f"Skipping file without an ID: {current_path}")
        return 0

    try:
        # Create the directory structure
        file_path = output_dir / current_path
        file_path.parent.mkdir(parents=True, exist_ok=True)

        # Check if file already exists and has the correct size
        if skip_existing and file_path.exists():
            existing_size = file_path.stat().st_size
            if existing_size == item.size:
                logger.debug(f" Skipping (already exists): {current_path}")
                return 1  # Count as downloaded
            else:
                logger.debug(
                    f"Re-downloading (size mismatch): {current_path} (local: {existing_size}, remote: {item.size})"
                )

        # Download the file content
        logger.debug(f"Downloading: {current_path} ({item.size} bytes)")

        content_stream = (
            await graph_client.drives.by_drive_id(drive_id)
            .items.by_drive_item_id(item.id)
            .content.get()
        )

        if content_stream:
            # Write to file
            with open(file_path, "wb") as f:
                f.write(content_stream)

            return 1
        else:
            logger.warning(f"No content for: {current_path}")
            return 0

    except Exception as e:
        if _is_auth_failure(e):
            raise
        logger.error(f"Failed to download {current_path}: {e}")
        return 0


async def download_drive(
    graph_client: GraphServiceClient,
    drive_id: str,
    output_dir: Path,
    max_concurrent: int = 20,
    skip_existing: bool = True,
) -> int:
    """Download every file in a drive into a folder named after the drive.

    Errors are logged rather than raised, except authentication failures,
    which abort the dump. A drive without a name is saved under its ID.

    Args:
        graph_client: Authenticated Graph client.
        drive_id: ID of the drive to download.
        output_dir: Directory in which the drive folder is created.
        max_concurrent: Maximum number of concurrent file downloads.
        skip_existing: Skip local files whose size already matches the
            remote size (resume mode).

    Returns:
        The number of files downloaded or skipped as already present; 0 if
        the drive is empty or could not be read.

    Raises:
        AuthenticationError: The Graph token is invalid or expired.
    """
    logger.info(f"Starting drive dump: {drive_id}")
    logger.info(f"Output directory: {output_dir}")
    logger.info(f"Max concurrent downloads: {max_concurrent}")
    if skip_existing:
        logger.info(" Resume mode: Skipping existing files")

    try:
        # Get drive info
        drive = await graph_client.drives.by_drive_id(drive_id).get()
        if drive is None:
            logger.error(f"Drive not found: {drive_id}")
            return 0
        logger.info(
            f"Drive: {drive.name} (Owner: {drive.owner.user.display_name if drive.owner and drive.owner.user else 'Unknown'})"
        )

        # Create output directory inside a folder named after the drive
        drive_folder = output_dir / (drive.name or drive_id)
        drive_folder.mkdir(parents=True, exist_ok=True)
        logger.info(f"Files will be saved to: {drive_folder}")

        # Start from root - use items with "root" as the item ID (with pagination)
        root_children: list[DriveItem] = await pagination.collect_all(
            graph_client.drives.by_drive_id(drive_id)
            .items.by_drive_item_id("root")
            .children
        )

        if not root_children:
            logger.info("Drive is empty")
            return 0

        # Create semaphore to limit concurrent downloads
        semaphore = asyncio.Semaphore(max_concurrent)

        total_downloaded = 0

        # Process all items in root concurrently
        tasks = [
            download_drive_item(
                graph_client, drive_id, item, drive_folder, "", semaphore, skip_existing
            )
            for item in root_children
        ]
        results = await asyncio.gather(*tasks, return_exceptions=True)
        total_downloaded += _sum_results(results, "drive root")

        logger.success(f"Download complete! Total files: {total_downloaded}")
        return total_downloaded

    except Exception as e:
        if _is_auth_failure(e):
            raise
        logger.error(f"Failed to download drive: {e}")
        return 0


def add_arguments(parser: argparse.ArgumentParser) -> None:
    """Register the `sharepoint download` arguments on a parser.

    Args:
        parser: Parser for the `download` subcommand.
    """
    parser.add_argument(
        "indices",
        nargs="?",
        default=None,
        help="Download items by index from the last search (e.g., 1, '1,3', '2-5').",
    )

    parser.add_argument(
        "-c",
        "--concurrency",
        type=int,
        default=20,
        help="Maximum number of concurrent downloads. Default: 20.",
    )

    parser.add_argument(
        "--no-resume",
        action="store_true",
        help="Disable resume mode. Re-download all files even if they already exist.",
    )



@handle_graph_errors
@require_scopes("Files.Read.All")
async def run_with_arguments(context: GraphContext, args: argparse.Namespace) -> int:
    """Download cached search results by index, or a whole drive.

    With `indices`, downloads those items from the last `sharepoint search`.
    Otherwise downloads the drive given by the global `--drive-id` into
    `--save` (default: the current directory).

    Args:
        context: Shared Graph runtime context.
        args: Parsed arguments (`indices`, `concurrency`, `no_resume`, and
            the global `drive_id` and `save`).

    Returns:
        0 if at least one file was downloaded, 1 otherwise.

    Raises:
        AuthenticationError: The Graph token is invalid or expired.
    """
    # If indices provided, download from cached search results
    if args.indices:
        return await _download_from_cache(context, args)

    # Otherwise, full drive download requires --drive-id
    drive_id = getattr(args, "drive_id", None)

    if not drive_id:
        logger.error(
            "Provide indices from last search (e.g., '1,3') or --drive-id for full drive dump."
        )
        return 1

    # Parse output directory from global --save/--output/-o argument
    output_dir = Path(args.save or Path().cwd()).resolve()

    # Get concurrency setting
    max_concurrent = getattr(args, "concurrency", 20)

    # Determine if we should skip existing files (resume mode)
    skip_existing = not getattr(args, "no_resume", False)

    # Download the drive
    downloaded_count = await download_drive(
        context.graph_client, drive_id, output_dir, max_concurrent, skip_existing
    )

    return 0 if downloaded_count > 0 else 1


async def _download_from_cache(context: GraphContext, args: argparse.Namespace) -> int:
    """Download items from the last cached SharePoint search by index.

    Files are written flat into `--save` (default: the current directory);
    existing files are skipped and counted as downloaded.

    Args:
        context: Shared Graph runtime context.
        args: Parsed arguments (`indices` such as "1,3-5", and `save`).

    Returns:
        0 if at least one file was downloaded, 1 otherwise.

    Raises:
        AuthenticationError: The Graph token is invalid or expired. The
            caller's `handle_graph_errors` converts raw auth `ODataError`s.
    """
    cached = cache.load_results(key="sharepoint", identity=context.identity_hash)
    if not cached:
        logger.error("No cached search results. Run a search first.")
        return 1

    indices = cache.parse_indices(args.indices, len(cached))
    if not indices:
        logger.error(
            f"Invalid indices: {args.indices} (cached results: 1-{len(cached)})"
        )
        return 1

    output_dir = Path(args.save if args.save else ".").resolve()
    output_dir.mkdir(parents=True, exist_ok=True)

    downloaded = 0
    failed = 0

    for idx in indices:
        item = cached[idx]
        name = item["name"]
        drive_id = item["drive_id"]
        item_id = item["item_id"]

        if not drive_id or not item_id:
            logger.warning(f"Missing drive/item ID for: {name}")
            failed += 1
            continue

        author = item.get("author", "?")
        created = item.get("created", "")
        size_bytes = item.get("size") or 0
        if size_bytes >= 1_048_576:
            size_str = f"{size_bytes / 1_048_576:.1f} MB"
        elif size_bytes >= 1024:
            size_str = f"{size_bytes / 1024:.1f} KB"
        else:
            size_str = f"{size_bytes} B"

        try:
            safe_filename = name.replace("/", "_").replace("..", "_")
            file_path = output_dir / safe_filename

            if file_path.exists():
                logger.debug(f"Already exists: {safe_filename}")
                downloaded += 1
                continue

            file_stream = (
                await context.graph_client.drives.by_drive_id(drive_id)
                .items.by_drive_item_id(item_id)
                .content.get()
            )

            if file_stream:
                with open(file_path, "wb") as f:
                    f.write(file_stream)
                logger.info(f"{name}  {author}  {size_str}  {created}")
                downloaded += 1
            else:
                logger.warning(f"Empty content: {name}")
                failed += 1

        except Exception as exc:
            if _is_auth_failure(exc):
                raise
            logger.error(f"Failed to download {name}: {exc}")
            failed += 1

    logger.info(f"Downloaded {downloaded} file(s) to: {output_dir}")
    if failed:
        logger.warning(f"Failed: {failed}")

    return 0 if downloaded > 0 else 1
