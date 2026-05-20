# msgraphx/modules/teams/show.py
#
# Two modes:
#   teams show N          — render a cached result with surrounding context
#   teams show --chat NAME — show the last N messages from a named chat
#
# Required delegated permission: Chat.Read

# Built-in imports
from __future__ import annotations

import argparse

# External library imports
from kiota_abstractions.base_request_configuration import RequestConfiguration
from loguru import logger
from msgraph.generated.chats.item.messages.messages_request_builder import (
    MessagesRequestBuilder,
)
from msgraph.generated.models.chat_message import ChatMessage
from msgraph.generated.models.o_data_errors.o_data_error import ODataError
from msgraph.generated.users.item.chats.chats_request_builder import ChatsRequestBuilder
from rich.console import Console

# Local library imports
from ...core.context import GraphContext
from ...utils.cache import load_results, parse_indices
from ...utils.errors import handle_graph_errors
from ...utils.pagination import GraphPaginator
from ._common import extract_body

_DEFAULT_CONTEXT = 4
_DEFAULT_LAST = 20


def add_arguments(parser: "argparse.ArgumentParser") -> None:
    parser.add_argument(
        "index",
        nargs="?",
        type=str,
        help="Index (or range, e.g. 1-3) from the last teams search to display.",
    )
    parser.add_argument(
        "--chat",
        dest="chat_name",
        type=str,
        default=None,
        metavar="NAME",
        help="Show the last N messages from the chat whose topic or member name matches NAME.",
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
        help=f"Number of messages to show in --chat mode (default: {_DEFAULT_LAST}).",
    )


def _sender(msg: ChatMessage) -> str:
    if msg.from_ and msg.from_.user:
        return msg.from_.user.display_name or "?"
    return "?"


def _render_window(
    segment: list[ChatMessage],
    target_id: str,
    chat_label: str,
    console: Console,
) -> None:
    console.rule(f"[dim]{chat_label}[/dim]", style="dim")
    for msg in segment:
        body = extract_body(msg)
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

    if args.chat_name:
        return await _show_chat(context, args.chat_name, args.last)

    if not args.index:
        logger.error("Provide an index or --chat NAME.")
        return 1

    return await _show_cached(context, args.index, args.context)


async def _show_chat(context: "GraphContext", name: str, last: int) -> int:
    """Find a chat by name and print its last N messages."""
    name_lower = name.lower()

    chat_params = ChatsRequestBuilder.ChatsRequestBuilderGetQueryParameters(
        top=50,
        select=["id", "topic"],
        expand=["members"],
    )
    chat_config = RequestConfiguration(query_parameters=chat_params)

    chat_id: str | None = None
    chat_label: str | None = None

    async for chat in GraphPaginator(context.graph_client.me.chats, chat_config):
        topic = chat.topic or ""
        members = chat.members or []
        member_names = [
            (getattr(m, "display_name", None) or "").lower() for m in members
        ]
        if name_lower in topic.lower() or any(
            name_lower in mn for mn in member_names
        ):
            chat_id = chat.id
            chat_label = topic or ", ".join(
                getattr(m, "display_name", "") or "" for m in members
                if (getattr(m, "display_name", None) or "").lower() != "me"
            ) or chat.id
            break

    if not chat_id:
        logger.error(f"No chat found matching {name!r}.")
        return 1

    logger.info(f"Showing last {last} messages from: {chat_label}")

    msg_params = MessagesRequestBuilder.MessagesRequestBuilderGetQueryParameters(
        top=min(last, 50),
        orderby=["createdDateTime desc"],
    )
    msg_config = RequestConfiguration(query_parameters=msg_params)

    collected: list[ChatMessage] = []
    try:
        async for msg in GraphPaginator(
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

    console = Console()
    console.rule(f"[dim]{chat_label}[/dim]", style="dim")
    for msg in reversed(collected):
        body = extract_body(msg)
        sender = _sender(msg)
        created = msg.created_date_time
        sent = created.strftime("%Y-%m-%d %H:%M") if created else ""
        console.print(f"     {body}  [dim]{sender}[/dim]  [cyan]{sent}[/cyan]")

    return 0


async def _show_cached(context: "GraphContext", index: str, ctx_n: int) -> int:
    """Show a cached search result with surrounding conversation context."""
    cached = load_results(key="teams")
    if not cached:
        logger.error("No cached teams results. Run 'teams chat' or 'teams channel' first.")
        return 1

    indices = parse_indices(index, len(cached))
    if not indices:
        logger.error(f"Invalid index: {index} (cached: 1-{len(cached)})")
        return 1

    console = Console()

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
            async for msg in GraphPaginator(
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

        _render_window(segment, message_id, chat_label, console)

    return 0
