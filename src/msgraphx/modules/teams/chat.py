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
from ...utils.roles import require_scopes


async def fetch_chats(context: GraphContext, top: int = 20) -> list[dict]:
    """Return recent chats for the current user as plain dicts.

    Raises on API error. Callers are responsible for handling exceptions.
    """
    chat_params = ChatsRequestBuilder.ChatsRequestBuilderGetQueryParameters(
        top=top,
        select=["id", "topic"],
        expand=["lastMessagePreview", "members"],
        orderby=["lastMessagePreview/createdDateTime desc"],
    )
    chat_config = RequestConfiguration(query_parameters=chat_params)

    rows: list[dict] = []
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

        rows.append({"id": chat.id, "label": label, "last_message": last_preview, "last_date": last_ts})
        if len(rows) >= top:
            break

    return rows


async def fetch_search(
    context: GraphContext,
    query: str = "*",
    from_addr: str | None = None,
    after: str | None = None,
    before: str | None = None,
) -> list[dict]:
    """Return Teams chat messages matching a search query as plain dicts.

    `after` and `before` accept ISO 8601 date strings.
    Raises on API error. Callers are responsible for handling exceptions.
    """
    parts: list[str] = []
    if query and query != "*":
        parts.append(query)
    if from_addr:
        parts.append(f"from:{from_addr}")
    if after:
        parts.append(f"sent>={after.split('T')[0]}")
    if before:
        parts.append(f"sent<={before.split('T')[0]}")

    query_string = " ".join(parts) if parts else "*"

    search_options = graph_search.SearchOptions(
        query_string=query_string,
        sort_by=None,
        page_size=25,
    )

    items: list[dict] = []
    async for _item in graph_search.search_entities(
        context.graph_client,
        entity_types=[EntityType.ChatMessage],
        options=search_options,
    ):
        msg = _item if isinstance(_item, ChatMessage) else None
        if msg is None:
            continue

        created = msg.created_date_time
        body = strip_html(msg.body.content) if msg.body and msg.body.content else ""
        sender = "?"
        sender_id = None
        if msg.from_ and msg.from_.user:
            sender = msg.from_.user.display_name or "?"
            sender_id = msg.from_.user.id

        sent = created.strftime("%Y-%m-%d") if created else ""
        preview = body[:120] + "…" if len(body) > 120 else body

        items.append({
            "message_id": msg.id,
            "chat_id": msg.chat_id,
            "chat_label": msg.chat_id,
            "body": body,
            "body_preview": preview,
            "web_url": msg.web_url,
            "importance": str(msg.importance).split(".")[-1].lower() if msg.importance else None,
            "sent_datetime": created.isoformat() if created else None,
            "sent": sent,
            "sender": sender,
            "sender_id": sender_id,
        })

    return items


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
    rows = await fetch_chats(context, top=20)

    if context.json_output:
        output.print_json(rows)
        return 0

    if context.ndjson_output:
        for row in rows:
            output.print_ndjson_item(row)
        return 0

    table = Table(show_header=True, header_style="bold", box=None, padding=(0, 1))
    table.add_column("#", style="dim", justify="right", width=4)
    table.add_column("Chat", min_width=30)
    table.add_column("Last message", style="dim", max_width=60)
    table.add_column("Date", style="cyan", width=10)

    for i, row in enumerate(rows, 1):
        table.add_row(str(i), row["label"], row["last_message"], row["last_date"])

    console.print("[bold]Recent chats[/bold]")
    console.rule()
    console.print(table)
    return 0


@handle_graph_errors
@require_scopes("Chat.Read", "ChannelMessage.Read.All")
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

    after_iso: str | None = None
    before_iso: str | None = None

    if args.after:
        try:
            after_iso = parse_date_string(args.after)
        except ValueError as e:
            logger.error(str(e))
            return 1

    if args.before:
        try:
            before_iso = parse_date_string(args.before)
        except ValueError as e:
            logger.error(str(e))
            return 1

    logger.info(f"Chat search: {args.query!r}")
    cached_items: list[dict] = []

    if not context.json_output and not context.ndjson_output:
        console.print("[bold]Teams chat search results[/bold]")
        console.rule()

    try:
        cached_items = await fetch_search(
            context,
            query=args.query or "*",
            from_addr=args.from_addr,
            after=after_iso,
            before=before_iso,
        )
        for count, item in enumerate(cached_items, 1):
            if not context.json_output and not context.ndjson_output:
                console.print(
                    f"  [dim]{count:>4}.[/dim]  "
                    f"{item['body_preview']}  [dim]{item['sender']}[/dim]  [cyan]{item['sent']}[/cyan]"
                )
            if context.ndjson_output:
                output.print_ndjson_item(item)

    except KeyboardInterrupt:
        if cached_items:
            logger.info(f"Interrupted. {len(cached_items)} result(s) cached.")
    finally:
        if cached_items:
            cache.save_results(cached_items, key="teams", identity=context.identity_hash)

    if not context.json_output:
        console.rule()

    if not cached_items:
        logger.info("No results found.")
    else:
        logger.success(f"{len(cached_items)} message(s) found.")

    if context.json_output:
        output.print_json(cached_items)
    # ndjson items streamed inline

    return 0
