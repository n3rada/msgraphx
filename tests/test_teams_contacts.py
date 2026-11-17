# tests/test_teams_contacts.py

from __future__ import annotations

import pytest
from msgraph.generated.models.aad_user_conversation_member import (
    AadUserConversationMember,
)
from msgraph.generated.models.chat import Chat
from msgraph.generated.models.chat_type import ChatType
from msgraph.generated.models.conversation_member import ConversationMember

from msgraphx.modules.teams.contacts import _is_one_on_one, _member_user_id


@pytest.mark.parametrize(
    "chat_type,expected",
    [
        (ChatType.OneOnOne, True),
        (ChatType.Group, False),
        (ChatType.Meeting, False),
        (None, False),
    ],
)
def test_is_one_on_one(chat_type: ChatType | None, expected: bool):
    assert _is_one_on_one(Chat(chat_type=chat_type)) is expected


def test_member_user_id_prefers_entra_user_id():
    member = AadUserConversationMember(id="membership-1", user_id="user-1")
    assert _member_user_id(member) == "user-1"


def test_member_user_id_falls_back_to_membership_id():
    assert _member_user_id(ConversationMember(id="membership-2")) == "membership-2"
    assert _member_user_id(AadUserConversationMember(id="membership-3")) == "membership-3"


def test_member_user_id_none_when_unknown():
    assert _member_user_id(ConversationMember()) is None
