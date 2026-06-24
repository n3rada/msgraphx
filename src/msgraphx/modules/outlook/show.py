# msgraphx/modules/outlook/show.py
#
# Fetch and render an email in the terminal from a cached mail search.
# Required delegated permission: Mail.Read
#
# MIME parsing uses Python's stdlib email + html.parser modules.
# Rendering uses rich (already a dependency).

# Built-in imports
from __future__ import annotations

import argparse
import email as _email_lib
from email.header import decode_header as _decode_header

# External library imports
from loguru import logger
from rich.console import Group
from rich.padding import Padding
from rich.panel import Panel
from rich.text import Text

# Local library imports
from ...core.context import GraphContext
from ...utils.cache import load_results, parse_indices
from ...utils.console import console
from ...utils.errors import handle_graph_errors
from ...utils.html import render_html

# ---------------------------------------------------------------------------
# MIME helpers
# ---------------------------------------------------------------------------


def _decode_value(value: str) -> str:
    parts = _decode_header(value or "")
    out: list[str] = []
    for data, charset in parts:
        if isinstance(data, bytes):
            out.append(data.decode(charset or "utf-8", errors="replace"))
        else:
            out.append(data)
    return "".join(out)


def _extract_mime(raw: bytes) -> dict:
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
            if payload is None:
                continue
            charset = part.get_content_charset() or "utf-8"
            text = payload.decode(charset, errors="replace")

            if ct == "text/plain" and not body_plain:
                body_plain = text
            elif ct == "text/html" and not body_html:
                body_html = text
    else:
        payload = msg.get_payload(decode=True)
        if payload:
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


def _render(parsed: dict) -> None:
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
    if body_raw and "<" in body_raw:
        body_renderable = render_html(body_raw)
    else:
        body_renderable = Text(body_raw or "(empty body)")

    # Attachments
    attach_text = Text()
    for name in parsed["attachments"]:
        attach_text.append(f"\n   {name}", style="yellow")

    # Compose panel using Group to mix Text and Markdown renderables
    parts = [header]
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


def add_arguments(parser: "argparse.ArgumentParser") -> None:
    parser.add_argument(
        "index",
        type=str,
        help="Index from the last mail search to display (e.g., 3).",
    )


@handle_graph_errors
async def run_with_arguments(
    context: "GraphContext", args: "argparse.Namespace"
) -> int:
    if context.is_app_only:
        logger.error("This module requires delegated authentication (user context).")
        return 1

    cached = load_results(key="mail")
    if not cached:
        logger.error("No cached mail search results. Run 'outlook search' first.")
        return 1

    indices = parse_indices(args.index, len(cached))
    if not indices:
        logger.error(f"Invalid index: {args.index} (cached: 1-{len(cached)})")
        return 1


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
        _render(parsed)

    return 0
