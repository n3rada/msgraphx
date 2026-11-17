# tests/test_teams_show.py

from __future__ import annotations

from datetime import datetime, timezone
from typing import Any, cast

from msgraph.generated.models.chat_message import ChatMessage
from msgraph.generated.models.chat_message_from_identity_set import (
    ChatMessageFromIdentitySet,
)
from msgraph.generated.models.identity import Identity
from msgraph.generated.models.item_body import ItemBody
from msgraph.generated.models.user import User

from msgraphx.core.context import GraphContext
from msgraphx.modules.teams.show import _message_record


def _context(user_id: str) -> GraphContext:
    return GraphContext(
        graph_client=cast(Any, None), is_app_only=False, cached_user=User(id=user_id)
    )


def _message(sender_id: str, sender_name: str) -> ChatMessage:
    return ChatMessage(
        id="m-1",
        created_date_time=datetime(2026, 1, 2, 3, 4, tzinfo=timezone.utc),
        body=ItemBody(content="<p>Hello <b>there</b></p>"),
        from_=ChatMessageFromIdentitySet(
            user=Identity(id=sender_id, display_name=sender_name)
        ),
    )


def test_message_record_fields():
    record = _message_record(_message("u-2", "Alice Smith"), _context("u-1"), "c-1", "Alice")

    assert record["chat_id"] == "c-1"
    assert record["chat_label"] == "Alice"
    assert record["message_id"] == "m-1"
    assert record["sender"] == "Alice Smith"
    assert record["sender_id"] == "u-2"
    assert record["sent_datetime"] == "2026-01-02T03:04:00+00:00"
    assert "<" not in record["body"] and "Hello" in record["body"]
    assert record["is_me"] is False


def test_message_record_flags_own_messages():
    record = _message_record(_message("u-1", "Me"), _context("u-1"), "c-1", "Alice")

    assert record["is_me"] is True
