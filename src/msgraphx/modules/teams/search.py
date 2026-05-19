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
import re
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

# Local library imports
from ...core.context import GraphContext
from ...utils.cache import save_results
from ...utils.dates import parse_date_string
from ...utils.errors import handle_graph_errors
from ...utils.pagination import GraphPaginator


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


def _strip_html(raw: str) -> str:
    body = re.sub(r"<[^>]+>", " ", raw).strip()
    return re.sub(r"\s{2,}", " ", body)


def _extract_body(msg: ChatMessage) -> str:
    if not msg.body:
        return ""
    raw = msg.body.content or ""
    if msg.body.content_type and str(msg.body.content_type) == "html":
        return _strip_html(raw)
    return raw.strip()


@handle_graph_errors
async def run_with_arguments(
    context: "GraphContext", args: "argparse.Namespace"
) -> int:
    if context.is_app_only:
        logger.error("This module requires delegated authentication (user context).")
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
    console.print("[bold]💬 Teams chat search results[/bold]")
    console.rule()

    count = 0
    cached_items: list[dict] = []

    chat_params = ChatsRequestBuilder.ChatsRequestBuilderGetQueryParameters(top=50)
    chat_config = RequestConfiguration(query_parameters=chat_params)

    try:
        async for chat in GraphPaginator(context.graph_client.me.chats, chat_config):
            chat_id = chat.id
            chat_label = chat.topic or chat_id

            msg_params = (
                MessagesRequestBuilder.MessagesRequestBuilderGetQueryParameters(top=50)
            )
            msg_config = RequestConfiguration(query_parameters=msg_params)

            try:
                async for msg in GraphPaginator(
                    context.graph_client.chats.by_chat_id(chat_id).messages, msg_config
                ):
                    if msg.deleted_date_time:
                        continue

                    created = msg.created_date_time
                    if created:
                        if created.tzinfo is None:
                            created = created.replace(tzinfo=timezone.utc)
                        if after_dt and created < after_dt:
                            continue
                        if before_dt and created > before_dt:
                            continue

                    body = _extract_body(msg)

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
                    preview = body[:120] + "…" if len(body) > 120 else body

                    console.print(
                        f"  [dim]{count:>4}.[/dim]  "
                        f"{preview}  [dim]{sender}[/dim]  [cyan]{sent}[/cyan]"
                    )

                    cached_items.append(
                        {
                            "message_id": msg.id,
                            "chat_id": chat_id,
                            "body": body,
                            "body_preview": preview,
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

    except KeyboardInterrupt:
        if count:
            logger.info(f"Interrupted — {count} result(s) cached.")
    finally:
        if cached_items:
            save_results(cached_items, key="teams")

    console.rule()
    if count == 0:
        logger.info("📭 No results found.")
    else:
        logger.info(f"✅ {count} message(s) found.")

    return 0
