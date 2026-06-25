# msgraphx/modules/outlook/download.py
#
# Download emails as .eml (MIME) files from a cached mail search.
#
# The Graph endpoint GET /me/messages/{id}/$value returns raw RFC 2822 MIME bytes.
# In the SDK: graph_client.me.messages.by_message_id(id).content.get()
#
# Tip: test the endpoint in Graph Explorer
# https://developer.microsoft.com/en-us/graph/graph-explorer

# Built-in imports
from __future__ import annotations

import argparse
from pathlib import Path

# External library imports
from loguru import logger

# Local library imports
from ...core.context import GraphContext
from ...utils import cache
from ...utils.errors import handle_graph_errors
from ...utils.roles import require_scopes

def add_arguments(parser: "argparse.ArgumentParser"):
    parser.add_argument(
        "indices",
        nargs="?",
        default=None,
        help="Download messages by index from the last mail search (e.g., 1, '1,3', '2-5').",
    )

@handle_graph_errors
@require_scopes("Mail.Read")
async def run_with_arguments(
    context: "GraphContext", args: "argparse.Namespace"
) -> int:
    if context.is_app_only:
        logger.error("This module requires delegated authentication (user context).")
        return 1

    cached = cache.load_results(key="mail", identity=context.identity_hash)
    if not cached:
        logger.error("No cached mail search results. Run 'outlook search' first.")
        return 1

    if not args.indices:
        logger.error("Provide indices to download (e.g., '1', '1,3', '2-5').")
        return 1

    indices = cache.parse_indices(args.indices, len(cached))
    if not indices:
        logger.error(f"Invalid indices: {args.indices} (cached: 1-{len(cached)})")
        return 1

    output_dir = Path(args.save if args.save else ".").resolve()
    output_dir.mkdir(parents=True, exist_ok=True)

    downloaded = 0
    failed = 0

    for idx in indices:
        item = cached[idx]
        message_id = item.get("message_id")
        subject = item.get("subject") or "no_subject"
        from_addr = item.get("from_address", "?")
        received = item.get("received", "")

        if not message_id:
            logger.warning(f"Missing message ID for: {subject}")
            failed += 1
            continue

        # Sanitize subject for use as filename
        safe_name = "".join(
            c if c.isalnum() or c in " ._-" else "_" for c in subject
        ).strip()
        safe_name = safe_name[:120]
        file_path = output_dir / f"{safe_name}.eml"

        if file_path.exists():
            logger.debug(f"Already exists: {file_path.name}")
            downloaded += 1
            continue

        try:
            # GET /me/messages/{id}/$value: raw MIME (RFC 2822) bytes
            mime_content = await context.graph_client.me.messages.by_message_id(
                message_id
            ).content.get()

            if mime_content:
                file_path.write_bytes(mime_content)
                logger.info(f"{subject}  {from_addr}  {received}")
                downloaded += 1
            else:
                logger.warning(f"Empty content: {subject}")
                failed += 1

        except Exception as exc:
            logger.error(f"Failed to download '{subject}': {exc}")
            failed += 1

    logger.info(f"Downloaded {downloaded} message(s) to: {output_dir}")
    if failed:
        logger.warning(f"Failed: {failed}")

    return 0 if downloaded > 0 else 1
