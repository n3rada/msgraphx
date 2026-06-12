# msgraphx/modules/teams/show.py
#
# Two modes, auto-detected from the positional argument:
#   teams show 3          — cached result with surrounding context (numeric index)
#   teams show alice      — last N messages from the chat matching that name
#
# Required delegated permission: Chat.Read

# Built-in imports
from __future__ import annotations

import argparse
import re

# External library imports
from kiota_abstractions.base_request_configuration import RequestConfiguration
from loguru import logger
from msgraph.generated.chats.item.messages.messages_request_builder import (
    MessagesRequestBuilder,
)
from msgraph.generated.models.chat_message import ChatMessage
from msgraph.generated.models.o_data_errors.o_data_error import ODataError
from msgraph.generated.users.item.chats.chats_request_builder import ChatsRequestBuilder
from rich.align import Align
from rich.text import Text

# Local library imports
from ...core.context import GraphContext
from ...utils import cache, pagination
from ...utils.console import console
from ...utils.errors import handle_graph_errors
from ...utils.html import render_html, strip_html

_DEFAULT_CONTEXT = 4
_DEFAULT_LAST = 20


def add_arguments(parser: "argparse.ArgumentParser") -> None:
    parser.add_argument(
        "target",
        type=str,
        help="Cached index / range (e.g. 3, 1-3) or a name to browse (e.g. alice).",
    )
    parser.add_argument(
        "--context",
        type=int,
        default=_DEFAULT_CONTEXT,
        metavar="N",
        help=f"Context messages around a cached match (default: {_DEFAULT_CONTEXT}).",
    )
    parser.add_argument(
        "--last",
        type=int,
        default=_DEFAULT_LAST,
        metavar="N",
        help=f"Number of messages to show in browse mode (default: {_DEFAULT_LAST}).",
    )


def _sender(msg: ChatMessage) -> str:
    if msg.from_ and msg.from_.user:
        return msg.from_.user.display_name or "?"
    return "?"


def _sender_short(msg: ChatMessage) -> str:
    """Return just the first name of the sender."""
    full = _sender(msg)
    return full.split()[0] if full != "?" else "?"


def _is_me(msg: ChatMessage, context: "GraphContext") -> bool:
    """Check if the message was sent by the current user."""
    if not context.cached_user or not msg.from_ or not msg.from_.user:
        return False
    return msg.from_.user.id == context.cached_user.id


def _render_window(
    segment: list[ChatMessage],
    target_id: str,
    chat_label: str,
) -> None:
    console.rule(f"[dim]{chat_label}[/dim]", style="dim")
    for msg in segment:
        body = strip_html(msg.body.content) if msg.body and msg.body.content else ""
        sender = _sender(msg)
        created = msg.created_date_time
        sent = created.strftime("%Y-%m-%d %H:%M") if created else ""
        if msg.id == target_id:
            console.print(
                f"  [bold]▶[/bold]  {body}  "
                f"[dim]{sender}[/dim]  [cyan]{sent}[/cyan]"
            )
        else:
            console.print(f"     [dim]{body}  {sender}  {sent}[/dim]")


@handle_graph_errors
async def run_with_arguments(
    context: "GraphContext", args: "argparse.Namespace"
) -> int:
    if context.is_app_only:
        logger.error("This module requires delegated authentication (user context).")
        return 1

    # Auto-detect mode: numeric/range → cached, anything else → chat browse
    if re.match(r"^[\d][0-9,\- ]*$", args.target):
        return await _show_cached(context, args.target, args.context)
    return await _show_chat(context, args.target, args.last)


async def _show_chat(context: "GraphContext", name: str, last: int) -> int:
    """Find a chat by name and print its last N messages."""
    name_lower = name.lower()

    chat_params = ChatsRequestBuilder.ChatsRequestBuilderGetQueryParameters(
        top=10,
        select=["id", "topic"],
        expand=["members"],
    )
    chat_config = RequestConfiguration(query_parameters=chat_params)

    # Collect (score, chat_id, chat_label) for all matching chats.
    # Score: 3 = exact member name, 2 = member starts-with, 1 = topic/member substring.
    # Smaller chats (1:1, small groups) get a bonus to prefer personal conversations.
    matches: list[tuple[int, str, str]] = []

    async for chat in pagination.GraphPaginator(context.graph_client.me.chats, chat_config):
        topic = chat.topic or ""
        members = chat.members or []
        member_count = len(members)

        score = 0
        for m in members:
            mn = (getattr(m, "display_name", None) or "").lower()
            words = re.split(r"[\s\(\),]+", mn)
            logger.debug(f"  member: {mn!r}  words: {words}")
            if mn == name_lower or name_lower in words:
                score = max(score, 3)
            elif mn.startswith(name_lower) or any(
                w.startswith(name_lower) for w in words
            ):
                score = max(score, 2)
            elif name_lower in mn:
                score = max(score, 1)
        if name_lower in topic.lower() and score < 2:
            score = max(score, 1)

        # Prefer 1:1 and small group chats over large groups
        if score and member_count <= 3:
            score += 2
        elif score and member_count <= 5:
            score += 1

        logger.debug(
            f"chat {(topic or chat.id)!r}: score={score} members={member_count}"
        )

        if score:
            label = (
                topic
                or ", ".join(
                    getattr(m, "display_name", "") or ""
                    for m in members
                    if (getattr(m, "display_name", None) or "").lower() != "me"
                )
                or chat.id
            )
            matches.append((score, chat.id, label))

            # Exact match in a 1:1 chat - no need to keep fetching
            if score >= 5:
                break

    if not matches:
        logger.error(f"No chat found matching {name!r}.")
        return 1

    matches.sort(key=lambda x: x[0], reverse=True)
    best_score, chat_id, chat_label = matches[0]

    if len(matches) > 1 and matches[1][0] == best_score:
        logger.warning(
            f"Multiple chats match {name!r}; showing the first. "
            f"Be more specific to narrow it down."
        )
        for _, _, lbl in matches[:5]:
            logger.info(f"  - {lbl}")

    logger.info(f"Showing last {last} messages from: {chat_label}")

    msg_params = MessagesRequestBuilder.MessagesRequestBuilderGetQueryParameters(
        top=min(last, 50),
        orderby=["createdDateTime desc"],
    )
    msg_config = RequestConfiguration(query_parameters=msg_params)

    collected: list[ChatMessage] = []
    try:
        async for msg in pagination.GraphPaginator(
            context.graph_client.chats.by_chat_id(chat_id).messages,
            msg_config,
        ):
            if msg.deleted_date_time:
                continue
            collected.append(msg)
            if len(collected) >= last:
                break
    except ODataError as exc:
        logger.error(f"Failed to fetch messages: {exc}")
        return 1

    console.rule(f"[dim]{chat_label}[/dim]", style="dim")
    for msg in reversed(collected):
        body = strip_html(msg.body.content) if msg.body and msg.body.content else ""
        created = msg.created_date_time
        sent = created.strftime("%H:%M") if created else ""

        if _is_me(msg, context):
            # Own messages: right-aligned
            line = Text()
            line.append(body)
            line.append(f"  {sent}", style="dim")
            console.print(Align.right(line))
        else:
            # Other person: left-aligned with colored name
            name = _sender_short(msg)
            line = Text()
            line.append(f"  {name}", style="bold cyan")
            line.append(f" {sent}", style="dim")
            line.append(f"  {body}")
            console.print(line)

    return 0


async def _show_cached(context: "GraphContext", index: str, ctx_n: int) -> int:
    """Show a cached search result with surrounding conversation context."""
    cached = cache.load_results(key="teams")
    if not cached:
        logger.error(
            "No cached teams results. Run 'teams chat' or 'teams channel' first."
        )
        return 1

    indices = cache.parse_indices(index, len(cached))
    if not indices:
        logger.error(f"Invalid index: {index} (cached: 1-{len(cached)})")
        return 1


    for idx in indices:
        item = cached[idx]
        chat_id = item.get("chat_id") or ""
        message_id = item.get("message_id") or ""
        chat_label = item.get("chat_label") or chat_id

        if not chat_id or not message_id:
            logger.warning(f"Missing chat_id or message_id for cached item {idx + 1}.")
            continue

        msg_params = MessagesRequestBuilder.MessagesRequestBuilderGetQueryParameters(
            top=50,
            orderby=["createdDateTime desc"],
        )
        msg_config = RequestConfiguration(query_parameters=msg_params)

        collected: list[ChatMessage] = []
        target_idx: int | None = None

        try:
            async for msg in pagination.GraphPaginator(
                context.graph_client.chats.by_chat_id(chat_id).messages,
                msg_config,
            ):
                if msg.deleted_date_time:
                    continue
                collected.append(msg)
                if msg.id == message_id:
                    target_idx = len(collected) - 1
                if target_idx is not None and len(collected) >= target_idx + ctx_n + 1:
                    break
        except ODataError as exc:
            logger.error(f"Failed to fetch messages for chat {chat_label}: {exc}")
            continue

        if target_idx is None:
            logger.warning(f"Message {message_id} not found in chat {chat_label}.")
            continue

        win_start = max(0, target_idx - ctx_n)
        win_end = min(len(collected), target_idx + ctx_n + 1)
        segment = list(reversed(collected[win_start:win_end]))

        _render_window(segment, message_id, chat_label)

    return 0
