# msgraphx/modules/teams/chat.py
#
# Search personal chat messages (1:1 DMs and group chats) via the Microsoft
# Search API (POST /search/query, EntityType.ChatMessage).
#
# Required delegated permissions:
#   Chat.Read              (user-consentable)
#   ChannelMessage.Read.All  (admin consent required)
#
# For channel messages (Teams workspaces), use: msgraphx teams channel
#
# Tip: prototype queries in Graph Explorer
# https://developer.microsoft.com/en-us/graph/graph-explorer

# Built-in imports
from __future__ import annotations

import argparse

# External library imports
from loguru import logger
from msgraph.generated.models.chat_message import ChatMessage
from msgraph.generated.models.entity_type import EntityType
from rich.console import Console

# Local library imports
from ...core import graph_search
from ...core.context import GraphContext
from ...utils.cache import save_results
from ...utils.dates import parse_date_string
from ...utils.errors import handle_graph_errors
from ._common import extract_body


def add_arguments(parser: "argparse.ArgumentParser") -> None:
    parser.add_argument(
        "query",
        nargs="?",
        default="*",
        help="Keyword to search for. Defaults to '*' (all messages).",
    )

    parser.add_argument(
        "--from",
        dest="from_addr",
        type=str,
        default=None,
        help="Filter by sender display name (client-side).",
    )


@handle_graph_errors
async def run_with_arguments(
    context: "GraphContext", args: "argparse.Namespace"
) -> int:
    if context.is_app_only:
        logger.error("This module requires delegated authentication (user context).")
        return 1

    if not context.has_scope("ChannelMessage.Read.All"):
        logger.error(
            "❌ Missing required scope: ChannelMessage.Read.All (admin consent required)."
        )
        return 1

    query = args.query or "*"

    parts: list[str] = []

    if query != "*":
        parts.append(query)

    if args.from_addr:
        parts.append(f"from:{args.from_addr}")

    if args.after:
        try:
            iso = parse_date_string(args.after)
            parts.append(f"sent>={iso.split('T')[0]}")
        except ValueError as e:
            logger.error(str(e))
            return 1

    if args.before:
        try:
            iso = parse_date_string(args.before)
            parts.append(f"sent<={iso.split('T')[0]}")
        except ValueError as e:
            logger.error(str(e))
            return 1

    query_string = " ".join(parts) if parts else "*"
    logger.info(f"🔍 Chat search: {query_string!r}")

    search_options = graph_search.SearchOptions(
        query_string=query_string,
        sort_by=None,
        page_size=25,
    )

    count = 0
    cached_items: list[dict] = []

    console = Console()
    console.print("[bold]💬 Teams chat search results[/bold]")
    console.rule()

    try:
        async for _item in graph_search.search_entities(
            context.graph_client,
            entity_types=[EntityType.ChatMessage],
            options=search_options,
        ):
            msg = _item if isinstance(_item, ChatMessage) else None
            if msg is None:
                continue

            logger.trace(msg.__dict__)

            created = msg.created_date_time

            body = extract_body(msg)

            sender = "?"
            sender_id = None
            if msg.from_ and msg.from_.user:
                sender = msg.from_.user.display_name or "?"
                sender_id = msg.from_.user.id

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
                    "chat_id": msg.chat_id,
                    "chat_label": msg.chat_id,
                    "body": body,
                    "web_url": msg.web_url,
                    "importance": str(msg.importance) if msg.importance else None,
                    "sent_datetime": created.isoformat() if created else None,
                    "sent": sent,
                    "sender": sender,
                    "sender_id": sender_id,
                }
            )

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
        logger.success(f"✅ {count} message(s) found.")

    return 0
