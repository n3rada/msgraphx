# tests/test_groups_members.py

from __future__ import annotations

from unittest.mock import AsyncMock, MagicMock

from msgraph.generated.models.group import Group
from msgraph.generated.models.group_collection_response import GroupCollectionResponse

from msgraphx.core.context import GraphContext
from msgraphx.modules.groups import members


def _context(groups: list[Group]) -> GraphContext:
    client = MagicMock()
    client.groups.get = AsyncMock(return_value=GroupCollectionResponse(value=groups))
    return GraphContext(graph_client=client, is_app_only=True)


async def test_guid_is_returned_unchanged():
    context = _context([])
    guid = "0f6c7ee1-6f4b-4f0e-9a8d-3b1b0b5c9f10"

    assert await members._resolve_group_id(context, guid) == guid
    context.graph_client.groups.get.assert_not_awaited()


async def test_mail_resolves_to_id():
    context = _context([Group(id="g-1", display_name="HR", mail="hr@contoso.com")])

    assert await members._resolve_group_id(context, "hr@contoso.com") == "g-1"
    config = context.graph_client.groups.get.await_args.kwargs["request_configuration"]
    assert config.query_parameters.filter == "mail eq 'hr@contoso.com'"


async def test_mail_with_quote_is_escaped():
    context = _context([Group(id="g-1", mail="o'brien@contoso.com")])

    await members._resolve_group_id(context, "o'brien@contoso.com")

    config = context.graph_client.groups.get.await_args.kwargs["request_configuration"]
    assert config.query_parameters.filter == "mail eq 'o''brien@contoso.com'"


async def test_unknown_mail_returns_none():
    assert await members._resolve_group_id(_context([]), "nobody@contoso.com") is None


async def test_ambiguous_mail_returns_none():
    context = _context([Group(id="g-1"), Group(id="g-2")])

    assert await members._resolve_group_id(context, "dup@contoso.com") is None
