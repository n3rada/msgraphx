# msgraphx/modules/teams/_common.py
#
# Shared helpers for Teams modules.

# Built-in imports
from __future__ import annotations

# External library imports
from msgraph.generated.models.chat_message import ChatMessage

# Local library imports
from ...utils.html import strip_html


def extract_body(msg: ChatMessage) -> str:
    """Return the plain-text body of a chat message, stripping HTML when needed."""
    if not msg.body:
        return ""
    raw = msg.body.content or ""
    if msg.body.content_type and str(msg.body.content_type) == "html":
        return strip_html(raw)
    return raw.strip()
