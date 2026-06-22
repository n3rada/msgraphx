# msgraphx/modules/onedrive/download.py

# Built-in imports
from __future__ import annotations

import argparse
from pathlib import Path

# External library imports
from loguru import logger

# Local library imports
from ...core.context import GraphContext
from ...modules.sharepoint.download import download_drive, _download_from_cache as _sp_download_from_cache
from ...utils import cache
from ...utils.errors import handle_graph_errors


def add_arguments(parser: argparse.ArgumentParser) -> None:
    parser.add_argument(
        "indices",
        nargs="?",
        default=None,
        help="Download items by index from the last OneDrive search (e.g., 1, '1,3', '2-5').",
    )
    parser.add_argument(
        "-c",
        "--concurrency",
        type=int,
        default=20,
        help="Maximum concurrent downloads. Default: 20.",
    )
    parser.add_argument(
        "--no-resume",
        action="store_true",
        help="Re-download files even if they already exist.",
    )


@handle_graph_errors
async def run_with_arguments(context: GraphContext, args: argparse.Namespace) -> int:
    if args.indices:
        return await _download_from_cache(context, args)

    drive_id = getattr(args, "drive_id", None)
    if not drive_id:
        logger.error(
            "Provide indices from last search (e.g., '1,3') or --drive-id for full drive dump."
        )
        return 1

    output_dir = Path(args.save or Path().cwd()).resolve()
    max_concurrent = getattr(args, "concurrency", 20)
    skip_existing = not getattr(args, "no_resume", False)

    downloaded = await download_drive(
        context.graph_client, drive_id, output_dir, max_concurrent, skip_existing
    )
    return 0 if downloaded > 0 else 1


async def _download_from_cache(context: GraphContext, args: argparse.Namespace) -> int:
    cached = cache.load_results(key="onedrive")
    if not cached:
        logger.error("No cached OneDrive search results. Run 'onedrive search' first.")
        return 1

    indices = cache.parse_indices(args.indices, len(cached))
    if not indices:
        logger.error(f"Invalid indices: {args.indices} (cached results: 1-{len(cached)})")
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
                logger.info(f"{name}  {item.get('author', '?')}  {size_str}  {item.get('created', '')}")
                downloaded += 1
            else:
                logger.warning(f"Empty content: {name}")
                failed += 1

        except Exception as exc:
            logger.error(f"Failed to download {name}: {exc}")
            failed += 1

    logger.info(f"Downloaded {downloaded} file(s) to: {output_dir}")
    if failed:
        logger.warning(f"Failed: {failed}")

    return 0 if downloaded > 0 else 1
