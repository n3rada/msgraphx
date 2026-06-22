# msgraphx/modules/teams/channel.py
#
# Search Teams channel messages via the Microsoft Graph Search API.
# Channels are shared topic spaces inside a Team workspace. All team
# members can see them.  The Graph Search API searches across all channels
# the caller has access to in a single request.
#
# Required delegated permissions:
#   Chat.Read              (user-consentable)
#   ChannelMessage.Read.All  (admin consent required)
#
# For personal chat messages (1:1 / group DMs), use: msgraphx teams chat
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
from rich.table import Table

# Local library imports
from ...core import graph_search
from ...core.context import GraphContext
from ...utils import cache, output
from ...utils.console import console
from ...utils.dates import parse_date_string
from ...utils.errors import handle_graph_errors
from ...utils.html import strip_html


async def fetch_teams(context: GraphContext) -> list[dict]:
    """Return joined Teams for the current user as plain dicts.

    Raises on API error. Callers are responsible for handling exceptions.
    """
    response = await context.graph_client.me.joined_teams.get()
    teams = (response.value or []) if response else []
    return [
        {"id": t.id, "display_name": t.display_name, "description": t.description}
        for t in teams
    ]


async def fetch_search(
    context: GraphContext,
    query: str = "*",
    from_addr: str | None = None,
    after: str | None = None,
    before: str | None = None,
) -> list[dict]:
    """Return Teams channel messages matching a search query as plain dicts.

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

        body = strip_html(msg.body.content) if msg.body and msg.body.content else ""
        sender = "?"
        sender_id = None
        if msg.from_ and msg.from_.user:
            sender = msg.from_.user.display_name or "?"
            sender_id = msg.from_.user.id

        sent = (
            msg.created_date_time.strftime("%Y-%m-%d") if msg.created_date_time else ""
        )
        team_id = None
        channel_id = None
        if msg.channel_identity:
            team_id = msg.channel_identity.team_id
            channel_id = msg.channel_identity.channel_id

        preview = body[:120] + "…" if len(body) > 120 else body

        items.append({
            "message_id": msg.id,
            "chat_id": msg.chat_id,
            "team_id": team_id,
            "channel_id": channel_id,
            "body": body,
            "body_preview": preview,
            "web_url": msg.web_url,
            "importance": str(msg.importance).split(".")[-1].lower() if msg.importance else None,
            "sent_datetime": (
                msg.created_date_time.isoformat() if msg.created_date_time else None
            ),
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
        help="KQL search query. Omit to list joined teams.",
    )

    parser.add_argument(
        "--from",
        dest="from_addr",
        type=str,
        default=None,
        help="Filter by sender display name or UPN.",
    )


async def _list_teams(context: "GraphContext") -> int:
    rows = await fetch_teams(context)

    if context.json_output:
        output.print_json(rows)
        return 0

    if context.ndjson_output:
        for row in rows:
            output.print_ndjson_item(row)
        return 0

    table = Table(show_header=True, header_style="bold", box=None, padding=(0, 1))
    table.add_column("#", style="dim", justify="right", width=4)
    table.add_column("Team", min_width=30)
    table.add_column("Description", style="dim")

    for i, row in enumerate(rows, 1):
        table.add_row(str(i), row["display_name"] or "", row["description"] or "")

    console.print("[bold]Joined Teams[/bold]")
    console.rule()
    console.print(table)
    logger.success(f"{len(rows)} team(s) found.")
    return 0


@handle_graph_errors
async def run_with_arguments(
    context: "GraphContext", args: "argparse.Namespace"
) -> int:
    if context.is_app_only:
        logger.error("This module requires delegated authentication (user context).")
        return 1

    # No query and no filters → list joined teams overview (only needs Team.ReadBasic.All)
    if not args.query and not args.from_addr:
        return await _list_teams(context)

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

    logger.info(f"Channel search: {args.query!r}")
    cached_items: list[dict] = []

    if not context.json_output and not context.ndjson_output:
        console.print("[bold]Teams channel search results[/bold]")
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
