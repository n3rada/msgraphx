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
from html.parser import HTMLParser

# External library imports
from loguru import logger
from rich.console import Console
from rich.padding import Padding
from rich.panel import Panel
from rich.rule import Rule
from rich.text import Text

# Local library imports
from ...core.context import GraphContext
from ...utils.cache import load_results, parse_indices
from ...utils.errors import handle_graph_errors


# ---------------------------------------------------------------------------
# MIME helpers
# ---------------------------------------------------------------------------


class _StripHTML(HTMLParser):
    """Minimal HTML-to-plaintext converter using only stdlib."""

    _BLOCK_TAGS = {"p", "div", "tr", "li", "br", "h1", "h2", "h3", "h4", "h5", "h6"}
    _SKIP_TAGS = {"style", "script", "head"}

    def __init__(self) -> None:
        super().__init__()
        self._parts: list[str] = []
        self._skip = False

    def handle_starttag(self, tag: str, attrs) -> None:
        if tag in self._SKIP_TAGS:
            self._skip = True
        elif tag in self._BLOCK_TAGS:
            self._parts.append("\n")

    def handle_endtag(self, tag: str) -> None:
        if tag in self._SKIP_TAGS:
            self._skip = False

    def handle_data(self, data: str) -> None:
        if not self._skip:
            self._parts.append(data)

    def get_text(self) -> str:
        import re

        text = "".join(self._parts)
        # Collapse runs of blank lines to at most two
        return re.sub(r"\n{3,}", "\n\n", text).strip()


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

    body = body_plain or (_strip_html(body_html) if body_html else "")

    return {
        "subject": _decode_value(msg.get("Subject", "")),
        "from": _decode_value(msg.get("From", "")),
        "to": _decode_value(msg.get("To", "")),
        "cc": _decode_value(msg.get("Cc", "")),
        "date": msg.get("Date", ""),
        "body": body,
        "attachments": attachments,
    }


def _strip_html(html: str) -> str:
    p = _StripHTML()
    p.feed(html)
    return p.get_text()


# ---------------------------------------------------------------------------
# Rich rendering
# ---------------------------------------------------------------------------


def _render(parsed: dict, console: Console) -> None:
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

    # Body
    body_text = Text(parsed["body"] or "(empty body)")

    # Attachments
    attach_text = Text()
    for name in parsed["attachments"]:
        attach_text.append(f"\n  📎  {name}", style="yellow")

    # Compose inner content
    inner = Text.assemble(header)
    if parsed["body"]:
        inner.append("\n")
        inner.append_text(Text("─" * 2, style="dim"))
        inner.append("\n\n")
        inner.append_text(body_text)
    if parsed["attachments"]:
        inner.append("\n")
        inner.append_text(attach_text)

    console.print(
        Panel(
            Padding(inner, (0, 1)),
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

    console = Console()

    for idx in indices:
        item = cached[idx]
        message_id = item.get("message_id")
        subject = item.get("subject") or "(no subject)"

        if not message_id:
            logger.warning(f"Missing message ID for: {subject}")
            continue

        try:
            mime_bytes = (
                await context.graph_client.me.messages.by_message_id(
                    message_id
                ).content.get()
            )
        except Exception as exc:
            logger.error(f"Failed to fetch '{subject}': {exc}")
            continue

        if not mime_bytes:
            logger.warning(f"Empty content: {subject}")
            continue

        parsed = _extract_mime(mime_bytes)
        _render(parsed, console)

    return 0
