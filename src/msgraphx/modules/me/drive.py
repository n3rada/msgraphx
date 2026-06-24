# msgraphx/modules/me/drive.py
#
# Personal OneDrive operations: tree view and file upload.
#
#   me drive tree                    # full OneDrive tree from root
#   me drive tree --path Desktop     # tree rooted at a specific folder
#   me drive tree --depth 2          # limit recursion depth
#   me drive upload <file> [--path Desktop/subdir]
#
# Required delegated permissions: Files.ReadWrite
#
# Upload approach:
#   Files < 4 MB  → PUT /me/drive/root:/{dest}:/content  (single request)
#   Files >= 4 MB → createUploadSession + LargeFileUploadTask (msgraph_core)

from __future__ import annotations

# Built-in imports
import argparse
import asyncio
from io import BytesIO
from pathlib import Path

# External library imports
from loguru import logger
from msgraph.generated.drives.item.items.item.create_upload_session.create_upload_session_post_request_body import (
    CreateUploadSessionPostRequestBody,
)
from msgraph.generated.models.drive_item import DriveItem
from msgraph.generated.models.drive_item_uploadable_properties import DriveItemUploadableProperties
from msgraph_core.tasks import LargeFileUploadTask
from rich.tree import Tree

# Local library imports
from ...core.context import GraphContext
from ...utils import output
from ...utils.console import console
from ...utils.errors import handle_graph_errors
from ..sharepoint.download import download_drive, download_drive_item

_SMALL_FILE_LIMIT = 4 * 1024 * 1024  # 4 MB
_CHUNK_SIZE = 4 * 1024 * 1024        # 4 MB chunks for large uploads


def add_arguments(parser: argparse.ArgumentParser) -> None:
    sub = parser.add_subparsers(dest="drive_subcommand", required=True)

    tree_parser = sub.add_parser("tree", help="Print OneDrive folder structure as a tree.")
    tree_parser.add_argument(
        "--path",
        metavar="PATH",
        default="",
        help="Folder path to root the tree at (e.g. 'Desktop'). Defaults to root.",
    )
    tree_parser.add_argument(
        "--depth",
        metavar="N",
        type=int,
        default=3,
        help="Maximum recursion depth (default: 3).",
    )

    upload_parser = sub.add_parser("upload", help="Upload a local file to OneDrive.")
    upload_parser.add_argument(
        "file",
        metavar="FILE",
        help="Local file path to upload.",
    )
    upload_parser.add_argument(
        "--path",
        metavar="DEST",
        default="",
        help="Destination folder in OneDrive (e.g. 'Desktop'). Defaults to root.",
    )

    download_parser = sub.add_parser("download", help="Download a file or folder from OneDrive.")
    download_parser.add_argument(
        "--path",
        metavar="PATH",
        default="",
        help="OneDrive path to download (e.g. 'Desktop/report.pdf'). Omit for full drive dump.",
    )
    download_parser.add_argument(
        "--no-resume",
        action="store_true",
        help="Re-download files even if they already exist locally.",
    )
    download_parser.add_argument(
        "-c",
        "--concurrency",
        type=int,
        default=20,
        metavar="N",
        help="Max concurrent downloads (default: 20).",
    )


@handle_graph_errors
async def run_with_arguments(context: GraphContext, args: argparse.Namespace) -> int:
    drive = await context.graph_client.me.drive.get()
    if not drive or not drive.id:
        logger.error("Could not retrieve personal drive.")
        return 1

    sub = args.drive_subcommand

    if sub == "tree":
        return await _run_tree(context, drive.id, args)
    if sub == "upload":
        return await _run_upload(context, drive.id, args)
    if sub == "download":
        return await _run_download(context, drive.id, args)

    return 1


async def _list_children(context: GraphContext, drive_id: str, item_id: str) -> list:
    resp = (
        await context.graph_client.drives.by_drive_id(drive_id)
        .items.by_drive_item_id(item_id)
        .children.get()
    )
    return resp.value if resp and resp.value else []


async def _build_tree(
    context: GraphContext,
    drive_id: str,
    item_id: str,
    node: Tree,
    depth: int,
    max_depth: int,
) -> None:
    if depth > max_depth:
        node.add("[dim]…[/dim]")
        return

    children = await _list_children(context, drive_id, item_id)

    folders = [c for c in children if c.folder is not None]
    files = [c for c in children if c.folder is None]

    # Sort: folders first, then files, both alphabetically
    folders.sort(key=lambda c: (c.name or "").lower())
    files.sort(key=lambda c: (c.name or "").lower())

    for folder in folders:
        size_info = f"  [dim]{folder.folder.child_count} item(s)[/dim]" if folder.folder else ""
        branch = node.add(f"[bold cyan]{folder.name}/[/bold cyan]{size_info}")
        await _build_tree(context, drive_id, folder.id, branch, depth + 1, max_depth)

    for f in files:
        size_bytes = f.size or 0
        if size_bytes >= 1_048_576:
            size_str = f"{size_bytes / 1_048_576:.1f} MB"
        elif size_bytes >= 1024:
            size_str = f"{size_bytes / 1024:.1f} KB"
        else:
            size_str = f"{size_bytes} B"
        node.add(f"{f.name}  [dim]{size_str}[/dim]")


async def _run_tree(context: GraphContext, drive_id: str, args: argparse.Namespace) -> int:
    folder_path = (args.path or "").strip("/")
    max_depth = args.depth

    if folder_path:
        item_ref = f"root:/{folder_path}:"
        item = (
            await context.graph_client.drives.by_drive_id(drive_id)
            .items.by_drive_item_id(item_ref)
            .get()
        )
        if not item or not item.id:
            logger.error(f"Folder not found: {folder_path}")
            return 1
        root_id = item.id
        label = f"[bold]{folder_path}/[/bold]"
    else:
        root_item = await context.graph_client.drives.by_drive_id(drive_id).root.get()
        root_id = root_item.id
        label = "[bold]OneDrive/[/bold]"

    logger.info(f"Building tree (max depth {max_depth})…")
    tree = Tree(label)
    await _build_tree(context, drive_id, root_id, tree, depth=1, max_depth=max_depth)
    console.print(tree)
    return 0


async def _run_download(context: GraphContext, drive_id: str, args: argparse.Namespace) -> int:
    output_dir = Path(args.save if args.save else ".").resolve()
    output_dir.mkdir(parents=True, exist_ok=True)
    skip_existing = not getattr(args, "no_resume", False)
    concurrency = getattr(args, "concurrency", 20)
    folder_path = (args.path or "").strip("/")

    if not folder_path:
        count = await download_drive(
            context.graph_client, drive_id, output_dir, concurrency, skip_existing
        )
        return 0 if count > 0 else 1

    item_ref = f"root:/{folder_path}:"
    item = (
        await context.graph_client.drives.by_drive_id(drive_id)
        .items.by_drive_item_id(item_ref)
        .get()
    )
    if not item or not item.id:
        logger.error(f"Path not found: {folder_path}")
        return 1

    semaphore = asyncio.Semaphore(concurrency)
    parent = folder_path.rsplit("/", 1)[0] if "/" in folder_path else ""
    count = await download_drive_item(
        context.graph_client, drive_id, item, output_dir, parent, semaphore, skip_existing
    )
    logger.info(f"Downloaded {count} file(s) to: {output_dir}")
    return 0 if count > 0 else 1


async def _run_upload(context: GraphContext, drive_id: str, args: argparse.Namespace) -> int:
    local = Path(args.file)
    if not local.exists():
        logger.error(f"File not found: {local}")
        return 1
    if not local.is_file():
        logger.error(f"Not a file: {local}")
        return 1

    dest_folder = (args.path or "").strip("/")
    dest_path = f"{dest_folder}/{local.name}" if dest_folder else local.name
    size = local.stat().st_size

    logger.info(f"Uploading {local.name} ({size} bytes) → OneDrive:/{dest_path}")

    item_ref = f"root:/{dest_path}:"

    if size < _SMALL_FILE_LIMIT:
        data = local.read_bytes()
        result = (
            await context.graph_client.drives.by_drive_id(drive_id)
            .items.by_drive_item_id(item_ref)
            .content.put(data)
        )
        if result:
            logger.success(f"Uploaded: {result.web_url}")
        else:
            logger.error("Upload returned no result.")
            return 1
    else:
        result = await _chunked_upload(context, drive_id, item_ref, local, size)
        if not result:
            return 1

    if context.json_output:
        output.print_json({
            "name": result.name,
            "size": result.size,
            "web_url": result.web_url,
            "id": result.id,
        })

    return 0


async def _chunked_upload(
    context: GraphContext,
    drive_id: str,
    item_ref: str,
    local: Path,
    size: int,
):
    body = CreateUploadSessionPostRequestBody(
        item=DriveItemUploadableProperties(additional_data={"@microsoft.graph.conflictBehavior": "rename"})
    )
    session = (
        await context.graph_client.drives.by_drive_id(drive_id)
        .items.by_drive_item_id(item_ref)
        .create_upload_session.post(body)
    )
    if not session or not session.upload_url:
        logger.error("Failed to create upload session.")
        return None

    logger.info(f"Upload session created, sending {size // _CHUNK_SIZE + 1} chunk(s).")

    stream = BytesIO(local.read_bytes())
    task = LargeFileUploadTask(
        session,
        context.graph_client.request_adapter,
        stream,
        parsable_factory=DriveItem.create_from_discriminator_value,
        max_chunk_size=_CHUNK_SIZE,
    )
    result = await task.upload()
    item = result.item_response if result else None
    if item:
        logger.success(f"Uploaded: {getattr(item, 'web_url', '?')}")
    return item
