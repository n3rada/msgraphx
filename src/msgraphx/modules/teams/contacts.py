# msgraphx/modules/teams/contacts.py
#
# Build a communication graph from your Teams chats.
# Lists 1:1 chat partners and group chat participants.
#
# GET /me/chats?$expand=members gives all chats + their members in one call.
# For message frequency, use GET /me/chats/{id}/messages (expensive, opt-in via --count).
#
# Tip: prototype in Graph Explorer
# https://developer.microsoft.com/en-us/graph/graph-explorer

# Built-in imports
from __future__ import annotations

import argparse
import contextlib
import json
from collections import Counter
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

# External library imports
from kiota_abstractions.base_request_configuration import RequestConfiguration
from loguru import logger
from msgraph.generated.chats.item.messages.messages_request_builder import (
    MessagesRequestBuilder,
)
from msgraph.generated.models.chat import Chat
from msgraph.generated.models.chat_message_type import ChatMessageType
from msgraph.generated.models.chat_type import ChatType
from msgraph.generated.models.conversation_member import ConversationMember
from msgraph.generated.models.o_data_errors.o_data_error import ODataError
from msgraph.generated.users.item.chats.chats_request_builder import ChatsRequestBuilder
from rich.console import Group
from rich.live import Live
from rich.table import Table

# Local library imports
from ...core.context import GraphContext
from ...utils import output
from ...utils.console import console
from ...utils.dates import parse_date_string
from ...utils.errors import handle_graph_errors
from ...utils.pagination import GraphPaginator
from ...utils.roles import require_scopes

_MESSAGES_PAGE_SIZE = 50  # Graph maximum $top for /chats/{id}/messages

def _is_one_on_one(chat: Chat) -> bool:
    """Return whether the chat is a 1:1 chat (not a group or meeting chat)."""
    return chat.chat_type == ChatType.OneOnOne

def _member_user_id(member: ConversationMember) -> str | None:
    """Return the member's Entra user ID, falling back to the membership ID.

    Args:
        member: A chat member; `AadUserConversationMember` carries `user_id`.

    Returns:
        The user ID, the membership ID if there is no user ID, or `None`.
    """
    return getattr(member, "user_id", None) or member.id

async def _count_messages(
    context: GraphContext,
    chat_id: str,
    after_cutoff: datetime | None = None,
) -> int | None:
    """Count the user-authored messages in a chat.

    Pages through the chat's messages newest first, skipping deleted
    messages and system events, and stops at `after_cutoff`.

    Args:
        context: Shared Graph runtime context (delegated).
        chat_id: ID of the chat.
        after_cutoff: Only count messages created at or after this time.

    Returns:
        The message count, or `None` if the messages could not be read.
    """
    msg_params = MessagesRequestBuilder.MessagesRequestBuilderGetQueryParameters(
        top=_MESSAGES_PAGE_SIZE,
        orderby=["createdDateTime desc"],
    )
    msg_config = RequestConfiguration(query_parameters=msg_params)

    count = 0
    try:
        async for msg in GraphPaginator(
            context.graph_client.chats.by_chat_id(chat_id).messages, msg_config
        ):
            created = msg.created_date_time
            if after_cutoff and created:
                if created.tzinfo is None:
                    created = created.replace(tzinfo=timezone.utc)
                if created < after_cutoff:
                    break
            if msg.deleted_date_time or msg.message_type != ChatMessageType.Message:
                continue
            count += 1
    except ODataError as exc:
        logger.warning(f"Could not count messages in chat {chat_id}: {exc}")
        return None
    return count

def add_arguments(parser: argparse.ArgumentParser) -> None:
    """Register the `teams contacts` options on a parser.

    Args:
        parser: Parser for the `teams contacts` subcommand.
    """
    parser.set_defaults(uses_time_bounds=True)
    parser.add_argument(
        "--top",
        "-n",
        type=int,
        default=30,
        metavar="N",
        help="Number of top contacts to display (default: 30). Use 0 for all.",
    )

    parser.add_argument(
        "--count",
        action="store_true",
        help=(
            "Also count messages per chat and per contact (slower: pages "
            "through every chat's messages)."
        ),
    )

@handle_graph_errors
@require_scopes("Chat.Read")
async def run_with_arguments(
    context: GraphContext, args: argparse.Namespace
) -> int:
    """Rank the user's Teams contacts by number of shared chats.

    Pages through the user's chats, newest activity first, stopping at the
    `--after` bound, and counts 1:1 partners and group chat participants
    separately with live tables. With `--count`, also counts the messages
    in each chat (since `--after`) and totals them per contact. Optionally
    writes the full ranking to `args.save`. Under `--json` or `--ndjson` the
    live tables are suppressed and the full ranking is printed as one JSON
    document or one NDJSON line.

    Args:
        context: Shared Graph runtime context; must be delegated.
        args: Parsed arguments: `top`, `after`, `count`, `save`.

    Returns:
        0 on success, 1 for app-only auth, a missing `Chat.Read`/
        `Chat.ReadBasic` scope, or an invalid `--after` date.

    Raises:
        AuthenticationError: The Graph token is invalid or expired.
    """
    if context.is_app_only:
        logger.error("This module requires delegated authentication (user context).")
        return 1

    if not (context.has_scope("Chat.Read") or context.has_scope("Chat.ReadBasic")):
        logger.error(
            "Missing required scope. Grant at least Chat.ReadBasic or Chat.Read."
        )
        return 1

    me = await context.graph_client.me.get()
    my_id = me.id if me else None

    logger.info("Fetching Teams chats")

    query_params = ChatsRequestBuilder.ChatsRequestBuilderGetQueryParameters(
        expand=["members", "lastMessagePreview"],
        orderby=["lastMessagePreview/createdDateTime desc"],
        top=50,
    )
    request_config = RequestConfiguration(query_parameters=query_params)

    after_cutoff: datetime | None = None
    if args.after:
        try:
            after_cutoff = datetime.fromisoformat(
                parse_date_string(args.after)
            ).replace(tzinfo=timezone.utc)
        except ValueError as e:
            logger.error(str(e))
            return 1

    dm_counter: Counter[str] = Counter()
    group_counter: Counter[str] = Counter()
    names: dict[str, str] = {}
    msg_counter: Counter[str] = Counter()

    one_on_one: list[dict[str, Any]] = []
    groups: list[dict[str, Any]] = []
    total = 0

    top_n = args.top if args.top > 0 else None
    machine_output = context.json_output or context.ndjson_output

    def _build_table(counter: Counter[str], title: str) -> Table:
        table = Table(title=title, show_header=True, header_style="bold")
        table.add_column("#", justify="right", style="dim")
        table.add_column("Chats", justify="right")
        if args.count:
            table.add_column("Messages", justify="right")
        table.add_column("Contact")
        for rank, (uid, cnt) in enumerate(counter.most_common(top_n), 1):
            if args.count:
                table.add_row(
                    str(rank), str(cnt), str(msg_counter[uid]), names.get(uid, uid)
                )
            else:
                table.add_row(str(rank), str(cnt), names.get(uid, uid))
        return table

    def _ranking(counter: Counter[str]) -> list[dict[str, Any]]:
        rows: list[dict[str, Any]] = []
        for uid, cnt in counter.most_common():
            row: dict[str, Any] = {
                "uid": uid,
                "name": names.get(uid, uid),
                "chat_count": cnt,
            }
            if args.count:
                row["message_count"] = msg_counter[uid]
            rows.append(row)
        return rows

    live_ctx = contextlib.nullcontext() if machine_output else Live(console=console, refresh_per_second=4)
    with live_ctx as live:
        async for chat in GraphPaginator(context.graph_client.me.chats, request_config):
            total += 1
            if live is not None:
                live.update(
                    Group(
                        _build_table(
                            dm_counter,
                            f"Top {args.top or len(dm_counter)} : Direct message partners ({total:,} scanned)",
                        ),
                        _build_table(
                            group_counter,
                            f"Top {args.top or len(group_counter)} : Group chat participants",
                        ),
                    )
                )

            is_dm = _is_one_on_one(chat)
            members = chat.members or []

            # Chats are ordered newest-first; stop once we pass the lower date bound
            if after_cutoff and chat.last_message_preview:
                last_ts = chat.last_message_preview.created_date_time
                if last_ts:
                    if last_ts.tzinfo is None:
                        last_ts = last_ts.replace(tzinfo=timezone.utc)
                    if last_ts < after_cutoff:
                        break

            message_count: int | None = None
            if args.count and chat.id:
                message_count = await _count_messages(context, chat.id, after_cutoff)

            for m in members:
                uid = _member_user_id(m)
                display = m.display_name or ""
                if uid and uid != my_id:
                    if display:
                        names[uid] = display
                    if is_dm:
                        dm_counter[uid] += 1
                    else:
                        group_counter[uid] += 1
                    if message_count:
                        msg_counter[uid] += message_count

            entry: dict[str, Any]
            if is_dm:
                other = next(
                    (m for m in members if _member_user_id(m) != my_id), None
                )
                if not other:
                    continue
                entry = {
                    "chat_id": chat.id,
                    "display_name": other.display_name or "",
                    "email": getattr(other, "email", "") or "",
                }
                target = one_on_one
            else:
                entry = {
                    "chat_id": chat.id,
                    "topic": chat.topic or "",
                    "member_count": len(members),
                }
                target = groups
            if args.count:
                entry["message_count"] = message_count
            target.append(entry)

    logger.info(
        f"{len(one_on_one)} 1:1 chats, {len(groups)} group chats, "
        f"{len(dm_counter) + len(group_counter)} unique contacts"
    )

    if args.save or machine_output:
        data = {
            "one_on_one": one_on_one,
            "groups": groups,
            "dm_ranking": _ranking(dm_counter),
            "group_ranking": _ranking(group_counter),
        }

        if args.save:
            save_path = Path(args.save)
            save_path.parent.mkdir(parents=True, exist_ok=True)
            with open(save_path, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=2, ensure_ascii=False)
            logger.info(f"Full results saved to: {save_path.absolute()}")

        if context.json_output:
            output.print_json(data)
        elif context.ndjson_output:
            output.print_ndjson_item(data)

    return 0
