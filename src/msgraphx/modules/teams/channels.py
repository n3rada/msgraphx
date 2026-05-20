# msgraphx/modules/teams/channels.py
#
# Search Teams channel messages via the Microsoft Graph Search API.
# Channels are shared topic spaces inside a Team workspace — all team
# members can see them.  The Graph Search API searches across all channels
# the caller has access to in a single request.
#
# Required delegated permissions:
#   Chat.Read          (user-consentable)
#   ChannelMessage.Read.All  (admin consent required)
#
# Exchange-like cap: from + size <= 1000.
# Sorting is not supported for ChatMessage.
#
# For personal chat messages (1:1 / group DMs), use: msgraphx teams search
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
        help="KQL search query. Defaults to '*' (all channel messages).",
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
    logger.info(f"🔍 Channel search query: {query_string}")
    logger.info("ℹ️  Requires Chat.Read + ChannelMessage.Read.All (admin consent).")

    search_options = graph_search.SearchOptions(
        query_string=query_string,
        sort_by=None,
        page_size=500,
        max_pages=2,
    )

    count = 0
    cached_items: list[dict] = []

    console = Console()
    console.print("[bold]📢 Teams channel search results[/bold]")
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

            body = extract_body(msg)
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

            team_id = None
            channel_id = None
            if msg.channel_identity:
                team_id = msg.channel_identity.team_id
                channel_id = msg.channel_identity.channel_id

            preview = body[:120] + "…" if len(body) > 120 else body

            console.print(
                f"  [dim]{count:>4}.[/dim]  "
                f"{preview}  [dim]{sender}[/dim]  [cyan]{sent}[/cyan]"
            )

            cached_items.append(
                {
                    "message_id": msg.id,
                    "chat_id": msg.chat_id,
                    "team_id": team_id,
                    "channel_id": channel_id,
                    "body": body,
                    "body_preview": preview,
                    "web_url": msg.web_url,
                    "importance": str(msg.importance) if msg.importance else None,
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
