# msgraphx/utils/html.py
#
# HTML-to-text helpers shared across modules.
#
#   strip_html(raw)    — inline mode: strips tags, collapses whitespace to a
#                        single space.  Used for chat messages displayed on one
#                        line.
#
#   html_to_text(raw)  — block mode: preserves paragraph / list structure by
#                        inserting newlines at block-level elements.  Used for
#                        mail bodies rendered in a panel.

# Built-in imports
from __future__ import annotations

import re
from html.parser import HTMLParser


def strip_html(raw: str) -> str:
    """Strip HTML tags and collapse whitespace to a single space (inline display)."""
    body = re.sub(r"<[^>]+>", " ", raw).strip()
    return re.sub(r"\s{2,}", " ", body)


class _BlockAwareParser(HTMLParser):
    _BLOCK_TAGS = {"p", "div", "tr", "li", "br", "h1", "h2", "h3", "h4", "h5", "h6"}
    _SKIP_TAGS = {"style", "script", "head"}

    def __init__(self) -> None:
        super().__init__()
        self._parts: list[str] = []
        self._skip = False

    def handle_starttag(self, tag: str, attrs) -> None:  # noqa: ARG002
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
        text = "".join(self._parts)
        return re.sub(r"\n{3,}", "\n\n", text).strip()


def html_to_text(raw: str) -> str:
    """Convert HTML to plain text, preserving block-level structure (multiline display)."""
    parser = _BlockAwareParser()
    parser.feed(raw)
    return parser.get_text()
