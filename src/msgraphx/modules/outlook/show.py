# msgraphx/modules/outlook/show.py
#
# Fetch and render an email in the terminal from a cached mail search.
#
# MIME parsing uses Python's stdlib email + html.parser modules.
# Rendering uses rich (already a dependency).

# Built-in imports
from __future__ import annotations

import argparse
import email as _email_lib
from collections.abc import Mapping
from email.header import decode_header as _decode_header
from typing import Any

# External library imports
from loguru import logger
from rich.console import Group, RenderableType
from rich.padding import Padding
from rich.panel import Panel
from rich.text import Text

# Local library imports
from ...core.context import GraphContext
from ...utils import output
from ...utils.cache import load_results, parse_indices
from ...utils.console import console
from ...utils.errors import handle_graph_errors
from ...utils.html import render_html
from ...utils.roles import require_scopes

# ---------------------------------------------------------------------------
# MIME helpers
# ---------------------------------------------------------------------------

def _decode_value(value: str) -> str:
    """Decode an RFC 2047 encoded MIME header value to plain text.

    Args:
        value: Raw header value, possibly containing `=?charset?...?=`
            encoded words.

    Returns:
        The decoded text; undecodable bytes are replaced.
    """
    parts = _decode_header(value or "")
    out: list[str] = []
    for data, charset in parts:
        if isinstance(data, bytes):
            out.append(data.decode(charset or "utf-8", errors="replace"))
        else:
            out.append(data)
    return "".join(out)

def _extract_mime(raw: bytes) -> dict[str, Any]:
    """Parse a raw RFC 2822 message into display fields.

    Args:
        raw: The MIME message bytes, as returned by `/messages/{id}/$value`.

    Returns:
        A dict with decoded `subject`, `from`, `to`, `cc` and `date`
        strings, `body` (the first text/plain part, else the first
        text/html part, else empty), and `attachments` (a list of
        filenames).
    """
    msg = _email_lib.message_from_bytes(raw)

    body_plain = ""
    body_html = ""
    attachments: list[str] = []

    if msg.is_multipart():
        for part in msg.walk():
            ct = part.get_content_type()
            cd = str(part.get_content_disposition() or "")
            filename = part.get_filename()

            if filename or "attachment" in cd:
                attachments.append(_decode_value(filename or "unnamed"))
                continue

            payload = part.get_payload(decode=True)
            if not isinstance(payload, bytes):
                continue
            charset = part.get_content_charset() or "utf-8"
            text = payload.decode(charset, errors="replace")

            if ct == "text/plain" and not body_plain:
                body_plain = text
            elif ct == "text/html" and not body_html:
                body_html = text
    else:
        payload = msg.get_payload(decode=True)
        if isinstance(payload, bytes) and payload:
            charset = msg.get_content_charset() or "utf-8"
            text = payload.decode(charset, errors="replace")
            if msg.get_content_type() == "text/html":
                body_html = text
            else:
                body_plain = text

    body = body_plain or (body_html if body_html else "")

    return {
        "subject": _decode_value(msg.get("Subject", "")),
        "from": _decode_value(msg.get("From", "")),
        "to": _decode_value(msg.get("To", "")),
        "cc": _decode_value(msg.get("Cc", "")),
        "date": msg.get("Date", ""),
        "body": body,
        "attachments": attachments,
    }

# ---------------------------------------------------------------------------
# Rich rendering
# ---------------------------------------------------------------------------

def _render(parsed: Mapping[str, Any]) -> None:
    """Print a parsed message as a Rich panel on the shared console.

    HTML bodies are converted to Markdown for display.

    Args:
        parsed: Fields produced by `_extract_mime`.
    """
    subject = parsed["subject"] or "(no subject)"

    # Header grid
    header = Text()
    for label, value in [
        ("From", parsed["from"]),
        ("To", parsed["to"]),
        ("CC", parsed["cc"]),
        ("Date", parsed["date"]),
    ]:
        if value:
            header.append(f"  {label:<6}", style="bold dim")
            header.append(f"  {value}\n")

    # Body: render HTML as Markdown for rich display, plain text as-is
    body_raw = parsed["body"]
    body_renderable: RenderableType
    if body_raw and "<" in body_raw:
        body_renderable = render_html(body_raw)
    else:
        body_renderable = Text(body_raw or "(empty body)")

    # Attachments
    attach_text = Text()
    for name in parsed["attachments"]:
        attach_text.append(f"\n   {name}", style="yellow")

    # Compose panel using Group to mix Text and Markdown renderables
    parts: list[RenderableType] = [header]
    if body_raw:
        parts.append(Text("─" * 2, style="dim"))
        parts.append(body_renderable)
    if parsed["attachments"]:
        parts.append(attach_text)

    console.print(
        Panel(
            Padding(Group(*parts), (0, 1)),
            title=f"[bold]{subject}[/bold]",
            title_align="left",
            border_style="blue",
        )
    )

# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def add_arguments(parser: argparse.ArgumentParser) -> None:
    """Register the `outlook show` options on a parser.

    Args:
        parser: Parser for the `outlook show` subcommand.
    """
    parser.add_argument(
        "index",
        type=str,
        help="Index from the last mail search to display (e.g., 3).",
    )

@handle_graph_errors
@require_scopes("Mail.Read")
async def run_with_arguments(
    context: GraphContext, args: argparse.Namespace
) -> int:
    """Fetch cached mail search results and render them in the terminal.

    Messages are picked by 1-based index (e.g. `"3"` or `"1,3-5"`) from the
    last `outlook search`. A message that fails to fetch is logged and
    skipped. Under `--json` or `--ndjson` the parsed messages (`index`,
    `message_id` and the `_extract_mime` fields) are printed instead of Rich
    panels. Warns first if the token lacks `Mail.Read`.

    Args:
        context: Shared Graph runtime context; must be delegated.
        args: Parsed arguments: `index`.

    Returns:
        0 after rendering the selected messages, 1 for app-only auth, an
        empty cache, or an invalid index.

    Raises:
        AuthenticationError: The Graph token is invalid or expired.
    """
    if context.is_app_only:
        logger.error("This module requires delegated authentication (user context).")
        return 1

    cached = load_results(key="mail", identity=context.identity_hash)
    if not cached:
        logger.error("No cached mail search results. Run 'outlook search' first.")
        return 1

    indices = parse_indices(args.index, len(cached))
    if not indices:
        logger.error(f"Invalid index: {args.index} (cached: 1-{len(cached)})")
        return 1

    results: list[dict[str, Any]] = []
    for idx in indices:
        item = cached[idx]
        message_id = item.get("message_id")
        subject = item.get("subject") or "(no subject)"

        if not message_id:
            logger.warning(f"Missing message ID for: {subject}")
            continue

        try:
            mime_bytes = await context.graph_client.me.messages.by_message_id(
                message_id
            ).content.get()
        except Exception as exc:
            logger.error(f"Failed to fetch '{subject}': {exc}")
            continue

        if not mime_bytes:
            logger.warning(f"Empty content: {subject}")
            continue

        parsed = _extract_mime(mime_bytes)
        if context.json_output:
            results.append({"index": idx + 1, "message_id": message_id, **parsed})
        elif context.ndjson_output:
            output.print_ndjson_item(
                {"index": idx + 1, "message_id": message_id, **parsed}
            )
        else:
            _render(parsed)

    if context.json_output:
        output.print_json(results)

    return 0
