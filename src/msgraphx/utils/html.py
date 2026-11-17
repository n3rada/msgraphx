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
from bs4 import Tag
from markdownify import MarkdownConverter
from rich.markdown import Markdown

# ---------------------------------------------------------------------------
# Inline mode
# ---------------------------------------------------------------------------


def strip_html(raw: str) -> str:
    """Convert HTML to a single line of plain text for inline display.

    Args:
        raw: HTML markup.

    Returns:
        The text with tags removed, entities decoded, and whitespace runs
        collapsed to single spaces.
    """
    body = re.sub(r"<[^>]+>", " ", raw).strip()
    body = unescape(body)
    return re.sub(r"\s{2,}", " ", body)


# ---------------------------------------------------------------------------
# Markdown conversion (for Rich terminal rendering)
# ---------------------------------------------------------------------------


class _TeamsMarkdownConverter(MarkdownConverter):
    """A markdownify converter that understands Microsoft Teams HTML tags."""

    # markdownify 1.x calls converters as convert_<tag>(el, text, parent_tags=...).

    def convert_at(  # noqa: ARG002
        self, el: Tag, text: str, parent_tags: set[str]
    ) -> str:
        """Render an `<at>` mention tag as a bold @mention."""
        return f"**@{text.strip()}**"

    def convert_attachment(  # noqa: ARG002
        self, el: Tag, text: str, parent_tags: set[str]
    ) -> str:
        """Render an `<attachment>` tag as a placeholder."""
        return " [attachment] "

    def convert_img(  # noqa: ARG002
        self, el: Tag, text: str, parent_tags: set[str]
    ) -> str:
        """Render an image as a placeholder showing its alt text."""
        alt = el.get("alt", "image")
        return f" [{alt}] "


def html_to_markdown(raw: str) -> str:
    """Convert HTML to Markdown, handling Teams-specific tags.

    Args:
        raw: HTML markup, typically a Teams or Outlook message body.

    Returns:
        Markdown text with `<span>` tags stripped, mentions in bold, and
        attachments and images replaced by placeholders.
    """
    return _TeamsMarkdownConverter(strip=["span"]).convert(raw).strip()


def render_html(raw: str) -> Markdown:
    """Convert HTML into a Rich Markdown renderable for terminal display.

    Args:
        raw: HTML markup.

    Returns:
        A Rich `Markdown` object built from `html_to_markdown(raw)`.
    """
    return Markdown(html_to_markdown(raw))
