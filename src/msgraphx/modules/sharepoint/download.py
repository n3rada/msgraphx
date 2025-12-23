# msgraphx/modules/sharepoint/download.py

# Built-in imports
from typing import TYPE_CHECKING
from pathlib import Path
import asyncio

# External library imports
from loguru import logger
from msgraph.generated.models.drive_item import DriveItem

# Local library imports
from msgraphx.utils.pagination import collect_all


if TYPE_CHECKING:
    import argparse
    from msgraphx.core.context import GraphContext


async def download_drive_item(
    graph_client: "GraphServiceClient",
    drive_id: str,
    item: DriveItem,
    output_dir: Path,
    base_path: str = "",
    semaphore: asyncio.Semaphore = None,
    skip_existing: bool = True,
) -> int:
    """
    Download a single drive item (file or folder).

    Args:
        graph_client: Authenticated Graph client
        drive_id: The drive ID
        item: The DriveItem to download
        output_dir: Base output directory
        base_path: Current path relative to drive root
        semaphore: Semaphore to limit concurrent downloads

    Returns:
        Number of files downloaded
    """
    downloaded_count = 0

    # Build the full path
    current_path = f"{base_path}/{item.name}" if base_path else item.name

    if item.folder:
        # It's a folder - recurse into it
        logger.debug(f"📁 Entering folder: {current_path}")

        try:
            # Get all children of this folder (with pagination)
            children = await collect_all(
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

                for result in results:
                    if isinstance(result, Exception):
                        logger.error(f"❌ Error in folder {current_path}: {result}")
                    else:
                        downloaded_count += result

        except Exception as e:
            logger.error(f"❌ Failed to list folder {current_path}: {e}")

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


async def _download_file(
    graph_client: "GraphServiceClient",
    drive_id: str,
    item: DriveItem,
    output_dir: Path,
    current_path: str,
    skip_existing: bool = True,
) -> int:
    """
    Helper function to download a single file.

    Args:
        graph_client: Authenticated Graph client
        drive_id: The drive ID
        item: The DriveItem to download
        output_dir: Base output directory
        current_path: Full path for the file
        skip_existing: Skip files that already exist with matching size

    Returns:
        1 if downloaded successfully, 0 otherwise
    """
    try:
        # Create the directory structure
        file_path = output_dir / current_path
        file_path.parent.mkdir(parents=True, exist_ok=True)

        # Check if file already exists and has the correct size
        if skip_existing and file_path.exists():
            existing_size = file_path.stat().st_size
            if existing_size == item.size:
                logger.debug(f"⏭️  Skipping (already exists): {current_path}")
                return 1  # Count as downloaded
            else:
                logger.debug(
                    f"🔄 Re-downloading (size mismatch): {current_path} (local: {existing_size}, remote: {item.size})"
                )

        # Download the file content
        logger.debug(f"📥 Downloading: {current_path} ({item.size} bytes)")

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
            logger.warning(f"⚠️ No content for: {current_path}")
            return 0

    except Exception as e:
        logger.error(f"❌ Failed to download {current_path}: {e}")
        return 0


async def download_drive(
    graph_client: "GraphServiceClient",
    drive_id: str,
    output_dir: Path,
    max_concurrent: int = 20,
    skip_existing: bool = True,
) -> int:
    """
    Download all files from a drive.

    Args:
        graph_client: Authenticated Graph client
        drive_id: The drive ID to download
        output_dir: Directory to save files
        max_concurrent: Maximum number of concurrent downloads
        skip_existing: Skip files that already exist with matching size

    Returns:
        Total number of files downloaded
    """
    logger.info(f"🚀 Starting drive dump: {drive_id}")
    logger.info(f"💾 Output directory: {output_dir}")
    logger.info(f"⚡ Max concurrent downloads: {max_concurrent}")
    if skip_existing:
        logger.info(f"⏭️  Resume mode: Skipping existing files")

    try:
        # Get drive info
        drive = await graph_client.drives.by_drive_id(drive_id).get()
        logger.info(
            f"📊 Drive: {drive.name} (Owner: {drive.owner.user.display_name if drive.owner and drive.owner.user else 'Unknown'})"
        )

        # Create output directory inside a folder named after the drive
        drive_folder = output_dir / drive.name
        drive_folder.mkdir(parents=True, exist_ok=True)
        logger.info(f"📂 Files will be saved to: {drive_folder}")

        # Start from root - use items with "root" as the item ID (with pagination)
        root_children = await collect_all(
            graph_client.drives.by_drive_id(drive_id)
            .items.by_drive_item_id("root")
            .children
        )

        if not root_children:
            logger.info("📭 Drive is empty")
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

        for result in results:
            if isinstance(result, Exception):
                logger.error(f"❌ Error during download: {result}")
            else:
                total_downloaded += result

        logger.success(f"🎉 Download complete! Total files: {total_downloaded}")
        return total_downloaded

    except Exception as e:
        logger.error(f"❌ Failed to download drive: {e}")
        return 0


def add_arguments(parser: "argparse.ArgumentParser"):
    parser.add_argument(
        "-o",
        "--output",
        type=str,
        default=Path().cwd(),
        help="Output directory for downloaded files. Defaults to the current working directory.",
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

    parser.add_argument(
        "--max-size",
        type=int,
        default=None,
        help="Skip files larger than this size (in MB). Default: no limit.",
    )

    parser.add_argument(
        "--include",
        type=str,
        default=None,
        help="Only download files matching this pattern (e.g., '*.pdf', '*.docx').",
    )

    parser.add_argument(
        "--exclude",
        type=str,
        default=None,
        help="Skip files matching this pattern (e.g., '*.tmp', '*.log').",
    )


async def run_with_arguments(
    context: "GraphContext", args: "argparse.Namespace"
) -> int:

    # Get drive_id from args
    drive_id = getattr(args, "drive_id", None)

    if not drive_id:
        logger.error("❌ Drive ID is required for download. Use --drive-id to specify.")
        return 1

    # Parse output directory
    output_dir = Path(args.output).resolve()

    # Get concurrency setting
    max_concurrent = getattr(args, "concurrency", 20)

    # Determine if we should skip existing files (resume mode)
    skip_existing = not getattr(args, "no_resume", False)

    # Download the drive
    downloaded_count = await download_drive(
        context.graph_client, drive_id, output_dir, max_concurrent, skip_existing
    )

    return 0 if downloaded_count > 0 else 1
