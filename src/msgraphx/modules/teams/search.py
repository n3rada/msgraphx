# msgraphx/modules/teams/search.py
#
# Search Teams messages via the Microsoft Search API (EntityType.ChatMessage).
# Required delegated permissions: Chat.Read, ChannelMessage.Read.All
#
# Exchange-like cap applies: from + size <= 1000.
# Sorting is not supported for ChatMessage.
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


def add_arguments(parser: "argparse.ArgumentParser") -> None:
    parser.add_argument(
        "query",
        nargs="?",
        default="*",
        help="KQL search query. Defaults to '*' (all messages).",
    )

    parser.add_argument(
        "--from",
        dest="from_addr",
        type=str,
        default=None,
        help="Filter by sender display name or UPN.",
    )


@handle_graph_errors
async def run_with_arguments(
    context: "GraphContext", args: "argparse.Namespace"
) -> int:
    if context.is_app_only:
        logger.error("This module requires delegated authentication (user context).")
        return 1

    parts: list[str] = []

    if args.query and args.query != "*":
        parts.append(args.query)

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
    logger.info(f"🔍 Search query: {query_string}")

    # ChatMessage has the same from+size<=1000 cap as Message
    search_options = graph_search.SearchOptions(
        query_string=query_string,
        sort_by=None,
        page_size=500,
        max_pages=2,
    )

    count = 0
    cached_items: list[dict] = []

    console = Console()
    console.print("[bold]💬 Teams search results[/bold]")
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

            count += 1

            sender = "?"
            sender_id = None
            if msg.from_ and msg.from_.user:
                sender = msg.from_.user.display_name or "?"
                sender_id = msg.from_.user.id

            sent = (
                msg.created_date_time.strftime("%Y-%m-%d")
                if msg.created_date_time
                else ""
            )

            # Determine context: channel or chat
            if msg.channel_identity:
                ctx = "[channel]"
                team_id = msg.channel_identity.team_id
                channel_id = msg.channel_identity.channel_id
                chat_id = None
            else:
                ctx = "[chat]"
                team_id = None
                channel_id = None
                chat_id = msg.chat_id

            # Strip HTML from body if present
            body = ""
            if msg.body:
                raw = msg.body.content or ""
                if msg.body.content_type and str(msg.body.content_type) == "html":
                    import re
                    body = re.sub(r"<[^>]+>", " ", raw).strip()
                    body = re.sub(r"\s{2,}", " ", body)
                else:
                    body = raw.strip()
            preview = body[:120] + "…" if len(body) > 120 else body

            console.print(
                f"  [dim]{count:>4}.[/dim]  [dim]{ctx}[/dim]  "
                f"{preview}  [dim]{sender}[/dim]  [cyan]{sent}[/cyan]"
            )

            cached_items.append(
                {
                    # Required for show/reply
                    "message_id": msg.id,
                    "chat_id": chat_id,
                    "team_id": team_id,
                    "channel_id": channel_id,
                    # Content
                    "body": body,
                    "body_preview": preview,
                    "web_url": msg.web_url,
                    "importance": str(msg.importance) if msg.importance else None,
                    # Dates
                    "sent_datetime": (
                        msg.created_date_time.isoformat()
                        if msg.created_date_time
                        else None
                    ),
                    "sent": sent,
                    # Sender
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

    if count == 0:
        logger.info("📭 No results found.")
        return 0

    return 0
