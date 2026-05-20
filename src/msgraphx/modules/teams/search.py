# msgraphx/modules/teams/search.py
#
# Search personal chat messages (1:1 DMs and group chats).
# Enumerates /me/chats, then pages through each chat's messages and
# filters client-side — no admin consent required.
#
# Required delegated permission: Chat.Read  (user-consentable)
#
# For channel messages (Teams workspaces), use: msgraphx teams channels
#
# Tip: prototype in Graph Explorer
# https://developer.microsoft.com/en-us/graph/graph-explorer

# Built-in imports
from __future__ import annotations

import argparse
import asyncio
from datetime import datetime, timezone

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
from rich.live import Live
from rich.text import Text

# Local library imports
from ...core.context import GraphContext
from ...utils.cache import save_results
from ...utils.dates import parse_date_string
from ...utils.errors import handle_graph_errors
from ...utils.pagination import GraphPaginator
from ._common import extract_body


def add_arguments(parser: "argparse.ArgumentParser") -> None:
    parser.add_argument(
        "query",
        nargs="?",
        default="*",
        help="Keyword to search for (case-insensitive substring match on message body). Defaults to '*' (all messages).",
    )

    parser.add_argument(
        "--from",
        dest="from_addr",
        type=str,
        default=None,
        help="Filter by sender display name or UPN.",
    )

    parser.add_argument(
        "--show",
        action="store_true",
        help="Display 4 messages of context before and after each match.",
    )


@handle_graph_errors
async def run_with_arguments(
    context: "GraphContext", args: "argparse.Namespace"
) -> int:
    if context.is_app_only:
        logger.error("This module requires delegated authentication (user context).")
        return 1

    if not (context.has_scope("Chat.Read") or context.has_scope("Chat.ReadBasic")):
        logger.error(
            "❌ Missing required scope. Grant at least Chat.ReadBasic or Chat.Read."
        )
        return 1

    query = args.query or "*"
    query_lower = query.lower() if query != "*" else None
    from_filter = args.from_addr.lower() if args.from_addr else None

    after_dt: datetime | None = None
    before_dt: datetime | None = None

    if args.after:
        try:
            iso = parse_date_string(args.after)
            dt = datetime.fromisoformat(iso)
            after_dt = dt.replace(tzinfo=timezone.utc) if dt.tzinfo is None else dt
        except ValueError as e:
            logger.error(str(e))
            return 1

    if args.before:
        try:
            iso = parse_date_string(args.before)
            dt = datetime.fromisoformat(iso)
            before_dt = dt.replace(tzinfo=timezone.utc) if dt.tzinfo is None else dt
        except ValueError as e:
            logger.error(str(e))
            return 1

    logger.info(
        f"🔍 Chat search: {'*' if query == '*' else repr(query)} "
        f"(personal chats — requires Chat.Read only)"
    )

    console = Console()
    count = 0
    chats_scanned = 0
    cached_items: list[dict] = []

    chat_params = ChatsRequestBuilder.ChatsRequestBuilderGetQueryParameters(
        top=50,
        select=["id", "topic"],
        expand=["lastMessagePreview"],
        orderby=["lastMessagePreview/createdDateTime desc"],
    )
    chat_config = RequestConfiguration(query_parameters=chat_params)

    with Live(console=console, refresh_per_second=4) as live:
        live.console.print("[bold]💬 Teams chat search results[/bold]")
        live.console.rule()

        try:
            async for chat in GraphPaginator(
                context.graph_client.me.chats, chat_config
            ):
                live.update(
                    Text(
                        f"  🔍 Scanning chat {chats_scanned}... {count} result(s) so far",
                        style="dim",
                    )
                )
                chat_id = chat.id
                chat_label = chat.topic or chat_id
                chats_scanned += 1

                # If chats are ordered newest-first, stop once we pass the lower bound
                if after_dt and chat.last_message_preview:
                    last_ts = chat.last_message_preview.created_date_time
                    if last_ts:
                        if last_ts.tzinfo is None:
                            last_ts = last_ts.replace(tzinfo=timezone.utc)
                        if last_ts < after_dt:
                            break

                msg_params = (
                    MessagesRequestBuilder.MessagesRequestBuilderGetQueryParameters(
                        top=50,
                        orderby=["createdDateTime desc"],
                    )
                )
                # createdDateTime only supports 'lt'; must be paired with $orderby on the same property.
                # Lower bound (after_dt) is enforced client-side below.
                if before_dt:
                    msg_params.filter = (
                        f"createdDateTime lt {before_dt.strftime('%Y-%m-%dT%H:%M:%SZ')}"
                    )
                msg_config = RequestConfiguration(query_parameters=msg_params)

                if args.show:
                    # Buffer all messages in this chat to enable context windows.
                    raw_msgs: list[ChatMessage] = []
                    try:
                        async for msg in GraphPaginator(
                            context.graph_client.chats.by_chat_id(chat_id).messages,
                            msg_config,
                        ):
                            if not msg.deleted_date_time:
                                raw_msgs.append(msg)
                    except ODataError as exc:
                        logger.debug(f"Skipping chat {chat_label}: {exc}")
                        continue

                    # Reverse to chronological order (oldest first) for context indexing.
                    raw_msgs.reverse()

                    for idx, msg in enumerate(raw_msgs):
                        created = msg.created_date_time
                        if created and created.tzinfo is None:
                            created = created.replace(tzinfo=timezone.utc)

                        if after_dt and created and created < after_dt:
                            continue

                        body = extract_body(msg)
                        if query_lower and query_lower not in body.lower():
                            continue

                        sender = "?"
                        sender_id = None
                        if msg.from_ and msg.from_.user:
                            sender = msg.from_.user.display_name or "?"
                            sender_id = msg.from_.user.id

                        if from_filter and from_filter not in sender.lower():
                            continue

                        count += 1
                        sent = created.strftime("%Y-%m-%d %H:%M") if created else ""

                        win_start = max(0, idx - 4)
                        win_end = min(len(raw_msgs), idx + 5)

                        live.console.rule(
                            f"[dim]  Match {count} · {chat_label}[/dim]", style="dim"
                        )
                        for ci in range(win_start, win_end):
                            cm = raw_msgs[ci]
                            if cm.deleted_date_time:
                                continue
                            cm_body = extract_body(cm)
                            cm_created = cm.created_date_time
                            cm_sent = (
                                cm_created.strftime("%Y-%m-%d %H:%M")
                                if cm_created
                                else ""
                            )
                            cm_sender = "?"
                            if cm.from_ and cm.from_.user:
                                cm_sender = cm.from_.user.display_name or "?"
                            if ci == idx:
                                live.console.print(
                                    f"  [bold]▶[/bold]  {cm_body}  "
                                    f"[dim]{cm_sender}[/dim]  [cyan]{cm_sent}[/cyan]"
                                )
                            else:
                                live.console.print(
                                    f"     [dim]{cm_body}  {cm_sender}  {cm_sent}[/dim]"
                                )

                        cached_items.append(
                            {
                                "message_id": msg.id,
                                "chat_id": chat_id,
                                "chat_label": chat_label,
                                "body": body,
                                "web_url": msg.web_url,
                                "importance": (
                                    str(msg.importance) if msg.importance else None
                                ),
                                "sent_datetime": (
                                    msg.created_date_time.isoformat()
                                    if msg.created_date_time
                                    else None
                                ),
                                "sent": sent,
                                "sender": sender,
                                "sender_id": sender_id,
                            }
                        )

                else:
                    # Streaming mode — break early once messages pass the lower bound.
                    try:
                        async for msg in GraphPaginator(
                            context.graph_client.chats.by_chat_id(chat_id).messages,
                            msg_config,
                        ):
                            if msg.deleted_date_time:
                                continue

                            created = msg.created_date_time
                            if created and created.tzinfo is None:
                                created = created.replace(tzinfo=timezone.utc)

                            # Messages are newest-first; stop once we pass the lower bound.
                            if after_dt and created and created < after_dt:
                                break

                            body = extract_body(msg)

                            if query_lower and query_lower not in body.lower():
                                continue

                            sender = "?"
                            sender_id = None
                            if msg.from_ and msg.from_.user:
                                sender = msg.from_.user.display_name or "?"
                                sender_id = msg.from_.user.id

                            if from_filter and from_filter not in sender.lower():
                                continue

                            count += 1
                            sent = created.strftime("%Y-%m-%d") if created else ""

                            live.console.print(
                                f"  [dim]{count:>4}.[/dim]  "
                                f"{body}  [dim]{sender}[/dim]  [cyan]{sent}[/cyan]"
                            )

                            cached_items.append(
                                {
                                    "message_id": msg.id,
                                    "chat_id": chat_id,
                                    "chat_label": chat_label,
                                    "body": body,
                                    "web_url": msg.web_url,
                                    "importance": (
                                        str(msg.importance) if msg.importance else None
                                    ),
                                    "sent_datetime": (
                                        msg.created_date_time.isoformat()
                                        if msg.created_date_time
                                        else None
                                    ),
                                    "sent": sent,
                                    "sender": sender,
                                    "sender_id": sender_id,
                                }
                            )

                    except ODataError as exc:
                        logger.debug(f"Skipping chat {chat_label}: {exc}")
                        continue

        except (KeyboardInterrupt, asyncio.CancelledError):
            if count:
                logger.info(f"Interrupted — {count} result(s) cached.")
            return 0
        finally:
            if cached_items:
                save_results(cached_items, key="teams")

        live.console.rule()

    if count == 0:
        logger.info("📭 No results found.")
    else:
        logger.success(f"✅ {count} message(s) found.")

    return 0
