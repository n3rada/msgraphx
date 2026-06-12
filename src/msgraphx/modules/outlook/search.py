# msgraphx/modules/outlook/search.py
#
# Search emails via the Microsoft Search API (EntityType.Message).
# Required delegated permission: Mail.Read (or Mail.ReadBasic)
#
# Exchange limitation: from + size <= 1000, so at most 2 pages of 500.
# Sorting is not supported for Message — the API ignores sortProperties.
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

    parts: list[str] = []

    if args.query and args.query != "*":
        parts.append(args.query)

    if args.from_addr:
        parts.append(f"from:{args.from_addr}")

    if args.subject:
        parts.append(f"subject:{args.subject}")

    if args.has_attachments:
        parts.append("hasAttachments:true")

    if args.after:
        try:
            iso = parse_date_string(args.after)
            parts.append(f"received>={iso.split('T')[0]}")
        except ValueError as e:
            logger.error(str(e))
            return 1

    if args.before:
        try:
            iso = parse_date_string(args.before)
            parts.append(f"received<={iso.split('T')[0]}")
        except ValueError as e:
            logger.error(str(e))
            return 1

    query_string = " ".join(parts) if parts else "*"
    logger.info(f"Search query: {query_string}")

    # Exchange caps: from + size <= 1000 → max 2 pages of 500
    # Sorting is not supported for Message (handled automatically in graph_search)
    search_options = graph_search.SearchOptions(
        query_string=query_string,
        sort_by=None,
        page_size=500,
        max_pages=2,
    )

    count = 0
    cached_items: list[dict] = []

    if not context.json_output:
        console.print("[bold]Mail search results[/bold]")
        console.rule()

    try:
        async for _item in graph_search.search_entities(
            context.graph_client,
            entity_types=[EntityType.Message],
            options=search_options,
        ):
            msg = _item if isinstance(_item, Message) else None
            if msg is None:
                continue

            logger.trace(msg.__dict__)

            count += 1

            from_addr = (
                msg.from_.email_address.address
                if msg.from_ and msg.from_.email_address
                else "?"
            )
            from_name = (
                msg.from_.email_address.name
                if msg.from_
                and msg.from_.email_address
                and msg.from_.email_address.name
                else ""
            )
            from_display = f"{from_name} <{from_addr}>" if from_name else from_addr

            received = (
                msg.received_date_time.strftime("%Y-%m-%d")
                if msg.received_date_time
                else ""
            )

            subject = msg.subject or "(no subject)"
            attach_flag = " [+att]" if msg.has_attachments else ""

            if not context.json_output:
                console.print(
                    f"  [dim]{count:>4}.[/dim]  {subject}{attach_flag}  "
                    f"[dim]{from_display}[/dim]  [cyan]{received}[/cyan]"
                )

            cached_items.append(
                {
                    # Required for download
                    "message_id": msg.id,
                    # Core
                    "subject": msg.subject,
                    "body_preview": msg.body_preview,
                    "web_link": msg.web_link,
                    "conversation_id": msg.conversation_id,
                    "internet_message_id": msg.internet_message_id,
                    "is_read": msg.is_read,
                    "importance": str(msg.importance).split(".")[-1].lower() if msg.importance else None,
                    "has_attachments": msg.has_attachments,
                    # Dates
                    "received_datetime": (
                        msg.received_date_time.isoformat()
                        if msg.received_date_time
                        else None
                    ),
                    "sent_datetime": (
                        msg.sent_date_time.isoformat() if msg.sent_date_time else None
                    ),
                    # Formatted date for display
                    "received": received,
                    # From
                    "from_address": from_addr,
                    "from_name": from_name,
                    # Recipients
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
                }
            )

    except KeyboardInterrupt:
        if count:
            logger.info(f"Interrupted — {count} result(s) cached.")
    finally:
        if cached_items:
            cache.save_results(cached_items, key="mail")

    if count == 0:
        logger.info("No results found.")
        if context.json_output:
            output.print_json([])
        return 0

    if context.json_output and cached_items:
        output.print_json(cached_items)

    return 0
