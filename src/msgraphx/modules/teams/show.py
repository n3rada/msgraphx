# msgraphx/modules/teams/show.py
#
# Two modes, auto-detected from the positional argument:
#   teams show 3          (cached result with surrounding context, numeric index)
#   teams show alice      (last N messages from the chat matching that name)

# Built-in imports
from __future__ import annotations

import argparse
import re
from collections.abc import Sequence
from typing import Any

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
from ...utils import cache, output, pagination
from ...utils.console import console
from ...utils.errors import handle_graph_errors
from ...utils.html import strip_html
from ...utils.roles import require_scopes

_DEFAULT_CONTEXT = 4
_DEFAULT_LAST = 20

def add_arguments(parser: argparse.ArgumentParser) -> None:
    """Register the `teams show` options on a parser.

    Args:
        parser: Parser for the `teams show` subcommand.
    """
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
    """Return the sender's display name, or `"?"` if unknown."""
    if msg.from_ and msg.from_.user:
        return msg.from_.user.display_name or "?"
    return "?"

def _sender_short(msg: ChatMessage) -> str:
    """Return the first word of the sender's display name, or `"?"`."""
    full = _sender(msg)
    return full.split()[0] if full != "?" else "?"

def _is_me(msg: ChatMessage, context: GraphContext) -> bool:
    """Return whether the message was sent by the context's cached user."""
    if not context.cached_user or not msg.from_ or not msg.from_.user:
        return False
    return msg.from_.user.id == context.cached_user.id

def _render_window(
    segment: Sequence[ChatMessage],
    target_id: str,
    chat_label: str,
) -> None:
    """Print a window of chat messages, highlighting the target message.

    Args:
        segment: Messages to print, oldest first.
        target_id: ID of the message to mark with `▶`.
        chat_label: Chat name shown in the header rule.
    """
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

def _message_record(
    msg: ChatMessage,
    context: GraphContext,
    chat_id: str,
    chat_label: str,
) -> dict[str, Any]:
    """Convert a chat message to a plain dict for JSON/NDJSON output.

    Args:
        msg: The chat message.
        context: Shared Graph runtime context, used to flag own messages.
        chat_id: ID of the chat the message belongs to.
        chat_label: Display label of the chat.

    Returns:
        A dict with `chat_id`, `chat_label`, `message_id`, `sender`,
        `sender_id`, `sent_datetime` (ISO 8601 or `None`), `body` (HTML
        stripped) and `is_me`.
    """
    created = msg.created_date_time
    return {
        "chat_id": chat_id,
        "chat_label": chat_label,
        "message_id": msg.id,
        "sender": _sender(msg),
        "sender_id": msg.from_.user.id if msg.from_ and msg.from_.user else None,
        "sent_datetime": created.isoformat() if created else None,
        "body": strip_html(msg.body.content) if msg.body and msg.body.content else "",
        "is_me": _is_me(msg, context),
    }

@handle_graph_errors
@require_scopes("Chat.Read")
async def run_with_arguments(
    context: GraphContext, args: argparse.Namespace
) -> int:
    """Show cached Teams search results in context, or browse a chat by name.

    A numeric target (e.g. `"3"` or `"1-3"`) selects cached results from the
    last `teams chat`/`teams channel` search and prints each with
    surrounding messages. Any other target is matched against chat members
    and topics, and the best match's latest messages are printed. Under
    `--json` or `--ndjson` the messages are printed as records (see
    `_message_record`; cached mode adds `index` and `is_target`), oldest
    first, instead of Rich output. Warns first if the token lacks
    `Chat.Read`.

    Args:
        context: Shared Graph runtime context; must be delegated.
        args: Parsed arguments: `target`, `context` (window size around a
            cached match) and `last` (messages to show when browsing).

    Returns:
        0 on success, 1 for app-only auth, an empty cache, an invalid index,
        no matching chat, or a failed message fetch in browse mode.

    Raises:
        AuthenticationError: The Graph token is invalid or expired.
    """
    if context.is_app_only:
        logger.error("This module requires delegated authentication (user context).")
        return 1

    # Auto-detect mode: numeric/range → cached, anything else → chat browse
    if re.match(r"^[\d][0-9,\- ]*$", args.target):
        return await _show_cached(context, args.target, args.context)
    return await _show_chat(context, args.target, args.last)

async def _show_chat(context: GraphContext, name: str, last: int) -> int:
    """Find the chat that best matches a name and print its latest messages.

    Chats are scored by how well a member name or the topic matches, with a
    bonus for small chats; scanning stops early on an exact 1:1 match.

    Args:
        context: Shared Graph runtime context (delegated).
        name: Person or topic to look for (case-insensitive).
        last: Number of most recent non-deleted messages to print.

    Returns:
        0 on success, 1 if no chat matches or fetching messages fails.

    Raises:
        ODataError: Listing the user's chats failed.
    """
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

    if context.json_output or context.ndjson_output:
        records = [
            _message_record(msg, context, chat_id, chat_label)
            for msg in reversed(collected)
        ]
        if context.json_output:
            output.print_json(records)
        else:
            for record in records:
                output.print_ndjson_item(record)
        return 0

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

async def _show_cached(context: GraphContext, index: str, ctx_n: int) -> int:
    """Print cached search results with surrounding conversation context.

    Items whose chat cannot be fetched, or whose message is not in the
    latest messages scanned, are logged and skipped.

    Args:
        context: Shared Graph runtime context (delegated).
        index: 1-based index spec into the cached results, e.g. `"1,3-5"`.
        ctx_n: Number of messages to show before and after each match.

    Returns:
        0 after processing the selected items, 1 for an empty cache or an
        invalid index.
    """
    cached = cache.load_results(key="teams", identity=context.identity_hash)
    if not cached:
        logger.error(
            "No cached teams results. Run 'teams chat' or 'teams channel' first."
        )
        return 1

    indices = cache.parse_indices(index, len(cached))
    if not indices:
        logger.error(f"Invalid index: {index} (cached: 1-{len(cached)})")
        return 1

    results: list[dict[str, Any]] = []
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

        if context.json_output or context.ndjson_output:
            for msg in segment:
                record = {
                    "index": idx + 1,
                    "is_target": msg.id == message_id,
                    **_message_record(msg, context, chat_id, chat_label),
                }
                if context.json_output:
                    results.append(record)
                else:
                    output.print_ndjson_item(record)
        else:
            _render_window(segment, message_id, chat_label)

    if context.json_output:
        output.print_json(results)

    return 0
