# msgraphx/modules/outlook/search.py
#
# Search emails via the Microsoft Search API (EntityType.Message).
# Required delegated permission: Mail.Read (or Mail.ReadBasic)
#
# Exchange limitation: from + size <= 1000, so at most 2 pages of 500.
# Sorting is not supported for Message; the API ignores sortProperties.
#
# Tip: prototype queries in Graph Explorer
# https://developer.microsoft.com/en-us/graph/graph-explorer

# Built-in imports
from __future__ import annotations

import argparse

# External library imports
from loguru import logger
from msgraph.generated.models.entity_type import EntityType
from msgraph.generated.models.message import Message

# Local library imports
from ...core import graph_search
from ...core.context import GraphContext
from ...utils import cache, output
from ...utils.console import console
from ...utils.dates import parse_date_string
from ...utils.errors import handle_graph_errors


async def fetch(
    context: GraphContext,
    query: str = "*",
    from_addr: str | None = None,
    subject: str | None = None,
    has_attachments: bool = False,
    after: str | None = None,
    before: str | None = None,
) -> list[dict]:
    """Return mail search results as plain dicts.

    `after` and `before` accept ISO 8601 date strings (YYYY-MM-DD or full datetime).
    Raises on API error. Callers are responsible for handling exceptions.
    """
    parts: list[str] = []

    if query and query != "*":
        parts.append(query)
    if from_addr:
        parts.append(f"from:{from_addr}")
    if subject:
        parts.append(f"subject:{subject}")
    if has_attachments:
        parts.append("hasAttachments:true")
    if after:
        parts.append(f"received>={after.split('T')[0]}")
    if before:
        parts.append(f"received<={before.split('T')[0]}")

    query_string = " ".join(parts) if parts else "*"

    search_options = graph_search.SearchOptions(
        query_string=query_string,
        sort_by=None,
        page_size=500,
        max_pages=2,
    )

    items: list[dict] = []
    async for _item in graph_search.search_entities(
        context.graph_client,
        entity_types=[EntityType.Message],
        options=search_options,
    ):
        msg = _item if isinstance(_item, Message) else None
        if msg is None:
            continue

        from_addr_val = (
            msg.from_.email_address.address
            if msg.from_ and msg.from_.email_address
            else "?"
        )
        from_name = (
            msg.from_.email_address.name
            if msg.from_ and msg.from_.email_address and msg.from_.email_address.name
            else ""
        )
        received = (
            msg.received_date_time.strftime("%Y-%m-%d") if msg.received_date_time else ""
        )

        items.append({
            "message_id": msg.id,
            "subject": msg.subject,
            "body_preview": msg.body_preview,
            "web_link": msg.web_link,
            "conversation_id": msg.conversation_id,
            "internet_message_id": msg.internet_message_id,
            "is_read": msg.is_read,
            "importance": str(msg.importance).split(".")[-1].lower() if msg.importance else None,
            "has_attachments": msg.has_attachments,
            "received_datetime": (
                msg.received_date_time.isoformat() if msg.received_date_time else None
            ),
            "sent_datetime": (
                msg.sent_date_time.isoformat() if msg.sent_date_time else None
            ),
            "received": received,
            "from_address": from_addr_val,
            "from_name": from_name,
            "to_recipients": [
                r.email_address.address
                for r in (msg.to_recipients or [])
                if r.email_address and r.email_address.address
            ],
            "cc_recipients": [
                r.email_address.address
                for r in (msg.cc_recipients or [])
                if r.email_address and r.email_address.address
            ],
        })

    return items


def add_arguments(parser: "argparse.ArgumentParser"):
    parser.add_argument(
        "query",
        nargs="?",
        default="*",
        help="KQL search query (e.g., 'password', 'from:alice@corp subject:vpn'). Defaults to '*'.",
    )

    parser.add_argument(
        "--from",
        dest="from_addr",
        type=str,
        default=None,
        help="Filter by sender address (KQL: from:<address>).",
    )

    parser.add_argument(
        "--subject",
        type=str,
        default=None,
        help="Filter by subject keyword (KQL: subject:<keyword>).",
    )

    parser.add_argument(
        "--has-attachments",
        action="store_true",
        help="Only show messages with attachments.",
    )


@handle_graph_errors
async def run_with_arguments(
    context: "GraphContext", args: "argparse.Namespace"
) -> int:
    if context.is_app_only:
        logger.error("This module requires delegated authentication (user context).")
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

    cached_items: list[dict] = []

    if not context.json_output and not context.ndjson_output:
        console.print("[bold]Mail search results[/bold]")
        console.rule()

    try:
        cached_items = await fetch(
            context,
            query=args.query or "*",
            from_addr=args.from_addr,
            subject=args.subject,
            has_attachments=args.has_attachments,
            after=after_iso,
            before=before_iso,
        )
        for count, item in enumerate(cached_items, 1):
            from_display = (
                f"{item['from_name']} <{item['from_address']}>"
                if item["from_name"]
                else item["from_address"]
            )
            subject = item["subject"] or "(no subject)"
            attach_flag = " [+att]" if item["has_attachments"] else ""
            if not context.json_output and not context.ndjson_output:
                console.print(
                    f"  [dim]{count:>4}.[/dim]  {subject}{attach_flag}  "
                    f"[dim]{from_display}[/dim]  [cyan]{item['received']}[/cyan]"
                )
            if context.ndjson_output:
                output.print_ndjson_item(item)

    except KeyboardInterrupt:
        if cached_items:
            logger.info(f"Interrupted. {len(cached_items)} result(s) cached.")
    finally:
        if cached_items:
            cache.save_results(cached_items, key="mail", identity=context.identity_hash)

    if not cached_items:
        logger.info("No results found.")
        if context.json_output:
            output.print_json([])
        return 0

    if context.json_output:
        output.print_json(cached_items)
    # ndjson items streamed inline

    return 0
