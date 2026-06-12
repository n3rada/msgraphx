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
from kiota_abstractions.base_request_configuration import RequestConfiguration
from loguru import logger
from msgraph.generated.models.chat_message import ChatMessage
from msgraph.generated.models.entity_type import EntityType
from msgraph.generated.users.item.chats.chats_request_builder import ChatsRequestBuilder
from rich.table import Table

# Local library imports
from ...core import graph_search
from ...core.context import GraphContext
from ...utils import cache, output
from ...utils.console import console
from ...utils.dates import parse_date_string
from ...utils.errors import handle_graph_errors
from ...utils.html import strip_html
from ...utils.pagination import GraphPaginator


def add_arguments(parser: "argparse.ArgumentParser") -> None:
    parser.set_defaults(uses_time_bounds=True)
    parser.add_argument(
        "query",
        nargs="?",
        default=None,
        help="Keyword to search for. Omit to list recent chats.",
    )

    parser.add_argument(
        "--from",
        dest="from_addr",
        type=str,
        default=None,
        help="Filter by sender display name (client-side).",
    )


async def _list_chats(context: "GraphContext") -> int:
    chat_params = ChatsRequestBuilder.ChatsRequestBuilderGetQueryParameters(
        top=20,
        select=["id", "topic"],
        expand=["lastMessagePreview", "members"],
        orderby=["lastMessagePreview/createdDateTime desc"],
    )
    chat_config = RequestConfiguration(query_parameters=chat_params)

    table = Table(show_header=True, header_style="bold", box=None, padding=(0, 1))
    table.add_column("#", style="dim", justify="right", width=4)
    table.add_column("Chat", min_width=30)
    table.add_column("Last message", style="dim", max_width=60)
    table.add_column("Date", style="cyan", width=10)

    count = 0
    chats_data: list[dict] = []

    async for chat in GraphPaginator(context.graph_client.me.chats, chat_config):
        members = chat.members or []
        label = (
            chat.topic
            or ", ".join(
                getattr(m, "display_name", "") or ""
                for m in members
                if (getattr(m, "display_name", None) or "").lower() != "me"
            )
            or chat.id
        )

        last_ts = ""
        last_preview = ""
        if chat.last_message_preview:
            ts = chat.last_message_preview.created_date_time
            if ts:
                last_ts = ts.strftime("%Y-%m-%d")
            body = chat.last_message_preview.body
            if body and body.content:
                raw = strip_html(body.content)
                last_preview = raw[:60] + "…" if len(raw) > 60 else raw

        count += 1
        chats_data.append({"id": chat.id, "label": label, "last_message": last_preview, "last_date": last_ts})
        table.add_row(str(count), label, last_preview, last_ts)
        if count >= 20:
            break

    if context.json_output:
        output.print_json(chats_data)
        return 0

    console.print("[bold]Recent chats[/bold]")
    console.rule()
    console.print(table)
    return 0


@handle_graph_errors
async def run_with_arguments(
    context: "GraphContext", args: "argparse.Namespace"
) -> int:
    if context.is_app_only:
        logger.error("This module requires delegated authentication (user context).")
        return 1

    # No query and no filters → list recent chats overview (only needs Chat.Read)
    if not args.query and not args.from_addr:
        return await _list_chats(context)

    if not context.has_scope("ChannelMessage.Read.All"):
        logger.error(
            "Missing required scope: ChannelMessage.Read.All (admin consent required)."
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
    logger.info(f"Chat search: {query_string!r}")

    search_options = graph_search.SearchOptions(
        query_string=query_string,
        sort_by=None,
        page_size=25,
    )

    count = 0
    cached_items: list[dict] = []

    if not context.json_output and not context.ndjson_output:
        console.print("[bold]Teams chat search results[/bold]")
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

            body = strip_html(msg.body.content) if msg.body and msg.body.content else ""

            sender = "?"
            sender_id = None
            if msg.from_ and msg.from_.user:
                sender = msg.from_.user.display_name or "?"
                sender_id = msg.from_.user.id

            count += 1
            sent = created.strftime("%Y-%m-%d") if created else ""
            preview = body[:120] + "…" if len(body) > 120 else body

            if not context.json_output and not context.ndjson_output:
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
                    "importance": str(msg.importance).split(".")[-1].lower() if msg.importance else None,
                    "sent_datetime": created.isoformat() if created else None,
                    "sent": sent,
                    "sender": sender,
                    "sender_id": sender_id,
                }
            )

            if context.ndjson_output:
                output.print_ndjson_item(cached_items[-1])

    except KeyboardInterrupt:
        if count:
            logger.info(f"Interrupted — {count} result(s) cached.")
    finally:
        if cached_items:
            cache.save_results(cached_items, key="teams")

    if not context.json_output:
        console.rule()

    if count == 0:
        logger.info("No results found.")
    else:
        logger.success(f"{count} message(s) found.")

    if context.json_output:
        output.print_json(cached_items)
    # ndjson items streamed inline

    return 0
