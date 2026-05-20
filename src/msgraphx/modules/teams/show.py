# msgraphx/modules/teams/show.py
#
# Render a cached Teams chat message with surrounding conversation context.
# Fetches messages from /chats/{id}/messages to build the context window.
#
# Required delegated permission: Chat.Read
# Run 'teams chat <query>' or 'teams channel <query>' first to populate the cache.

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
from rich.console import Console

# Local library imports
from ...core.context import GraphContext
from ...utils.cache import load_results, parse_indices
from ...utils.errors import handle_graph_errors
from ...utils.pagination import GraphPaginator
from ._common import extract_body

_DEFAULT_CONTEXT = 4


def add_arguments(parser: "argparse.ArgumentParser") -> None:
    parser.add_argument(
        "index",
        type=str,
        help="Index (or range, e.g. 1-3) from the last teams search to display.",
    )
    parser.add_argument(
        "--context",
        type=int,
        default=_DEFAULT_CONTEXT,
        metavar="N",
        help=f"Messages of context before and after the match (default: {_DEFAULT_CONTEXT}).",
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

    cached = load_results(key="teams")
    if not cached:
        logger.error("No cached teams search results. Run 'teams search' first.")
        return 1

    indices = parse_indices(args.index, len(cached))
    if not indices:
        logger.error(f"Invalid index: {args.index} (cached: 1-{len(cached)})")
        return 1

    ctx_n: int = args.context
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

        # Collect messages newest-first until we have the target + ctx_n older ones.
        # The ctx_n newer ones accumulate naturally before we reach the target.
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

        # collected is newest-first; slice the window and reverse for chronological display.
        win_start = max(0, target_idx - ctx_n)
        win_end = min(len(collected), target_idx + ctx_n + 1)
        segment = list(reversed(collected[win_start:win_end]))

        _render_window(segment, message_id, chat_label, console)

    return 0
