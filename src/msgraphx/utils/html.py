# msgraphx/utils/html.py
#
# Single source of truth for all HTML-to-text/terminal conversion.
#
#   strip_html(raw)        - Inline mode: one-line plain text (for search results,
#                            table cells, log lines).
#
#   html_to_markdown(raw)  - Converts HTML to Markdown via markdownify with custom
#                            handling for Teams-specific tags (<at>, <attachment>).
#
#   render_html(raw)       - Returns a Rich Markdown renderable for terminal display.

# Built-in imports
from __future__ import annotations

import re
from html import unescape

# External library imports
from markdownify import MarkdownConverter
from rich.markdown import Markdown

# ---------------------------------------------------------------------------
# Inline mode
# ---------------------------------------------------------------------------


def strip_html(raw: str) -> str:
    """Strip HTML tags, decode entities, collapse whitespace (inline display)."""
    body = re.sub(r"<[^>]+>", " ", raw).strip()
    body = unescape(body)
    return re.sub(r"\s{2,}", " ", body)


# ---------------------------------------------------------------------------
# Markdown conversion (for Rich terminal rendering)
# ---------------------------------------------------------------------------


class _TeamsMarkdownConverter(MarkdownConverter):
    """Custom markdownify converter for Microsoft Teams HTML."""

    def convert_at(self, el, text, convert_as_inline):  # noqa: ARG002
        """<at> mention tags become bold @mentions."""
        return f"**@{text.strip()}**"

    def convert_attachment(self, el, text, convert_as_inline):  # noqa: ARG002
        """<attachment> tags become a placeholder."""
        return " [attachment] "

    def convert_img(self, el, text, convert_as_inline):  # noqa: ARG002
        """Images become a placeholder (not useful in terminal)."""
        alt = el.get("alt", "image")
        return f" [{alt}] "


def html_to_markdown(raw: str) -> str:
    """Convert HTML to Markdown with Teams-specific tag handling."""
    return _TeamsMarkdownConverter(strip=["span"]).convert(raw).strip()


def render_html(raw: str) -> Markdown:
    """Return a Rich Markdown renderable from HTML content."""
    return Markdown(html_to_markdown(raw))
