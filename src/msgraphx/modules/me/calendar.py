# msgraphx/modules/me/calendar.py
#
# List calendar events for the current user.
# Meeting titles, attendees, and locations are OSINT goldmines for
# phishing lure construction and travel/schedule intelligence.
#
# Required delegated permissions:
#   Calendars.Read

# Built-in imports
from __future__ import annotations

import argparse
from datetime import datetime

# External library imports
from kiota_abstractions.base_request_configuration import RequestConfiguration
from loguru import logger
from msgraph.generated.users.item.calendar.events.events_request_builder import (
    EventsRequestBuilder,
)
from rich.table import Table

# Local library imports
from ...core.context import GraphContext
from ...utils import output
from ...utils.console import console
from ...utils.dates import parse_date_string
from ...utils.errors import handle_graph_errors
from ...utils.pagination import GraphPaginator


def add_arguments(parser: "argparse.ArgumentParser") -> None:
    parser.set_defaults(uses_time_bounds=True)
    parser.add_argument(
        "--top",
        "-n",
        type=int,
        default=50,
        metavar="N",
        help="Maximum number of events to return (default: 50).",
    )


async def fetch(
    context: GraphContext,
    top: int = 50,
    after: str | None = None,
    before: str | None = None,
) -> list[dict]:
    """Return calendar events for the current user as plain dicts.

    `after` and `before` accept ISO 8601 datetime strings for OData filtering.
    Raises on API error. Callers are responsible for handling exceptions.
    """
    filter_parts: list[str] = []
    if after:
        filter_parts.append(f"start/dateTime ge '{after}'")
    if before:
        filter_parts.append(f"end/dateTime le '{before}'")
    odata_filter = " and ".join(filter_parts) if filter_parts else None

    query_params = EventsRequestBuilder.EventsRequestBuilderGetQueryParameters(
        top=min(top, 999),
        filter=odata_filter,
        orderby=["start/dateTime desc"],
        select=["subject", "start", "end", "organizer", "attendees", "location",
                "isOnlineMeeting", "bodyPreview", "webLink"],
    )
    config = RequestConfiguration(query_parameters=query_params)

    rows = []
    async for event in GraphPaginator(context.graph_client.me.calendar.events, config):
        start = ""
        end = ""
        if event.start and event.start.date_time:
            try:
                dt = datetime.fromisoformat(event.start.date_time.rstrip("Z"))
                start = dt.strftime("%Y-%m-%d %H:%M")
            except ValueError:
                start = event.start.date_time[:16]
        if event.end and event.end.date_time:
            try:
                dt = datetime.fromisoformat(event.end.date_time.rstrip("Z"))
                end = dt.strftime("%H:%M")
            except ValueError:
                end = event.end.date_time[11:16]

        organizer = ""
        organizer_email = ""
        if event.organizer and event.organizer.email_address:
            organizer = event.organizer.email_address.name or ""
            organizer_email = event.organizer.email_address.address or ""

        location = ""
        if event.location and event.location.display_name:
            location = event.location.display_name

        attendees = [
            a.email_address.address
            for a in (event.attendees or [])
            if a.email_address and a.email_address.address
        ]

        rows.append({
            "id": event.id,
            "subject": event.subject,
            "start": start,
            "end": end,
            "organizer": organizer,
            "organizer_email": organizer_email,
            "location": location,
            "attendees": attendees,
            "is_online": event.is_online_meeting,
            "body_preview": event.body_preview,
            "web_link": event.web_link,
        })

        if len(rows) >= top:
            break

    return rows


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

    logger.info(f"Fetching calendar events (filter: after={after_iso!r} before={before_iso!r})")

    rows = await fetch(context, top=args.top, after=after_iso, before=before_iso)

    if not rows:
        logger.info("No events found.")
        if context.json_output:
            output.print_json([])
        return 0

    if context.json_output:
        output.print_json(rows)
        return 0

    if context.ndjson_output:
        for row in rows:
            output.print_ndjson_item(row)
        return 0

    table = Table(show_header=True, header_style="bold", box=None, padding=(0, 1))
    table.add_column("#", style="dim", justify="right", width=4)
    table.add_column("Subject", min_width=35)
    table.add_column("Start", style="cyan", width=16)
    table.add_column("Organizer", style="dim", min_width=20)
    table.add_column("Location", style="dim", max_width=25)

    for i, row in enumerate(rows, 1):
        subj = (row["subject"] or "(no subject)")[:60]
        table.add_row(str(i), subj, row["start"], row["organizer"], row["location"])

    console.print("[bold]Calendar events[/bold]")
    console.rule()
    console.print(table)
    logger.success(f"{len(rows)} event(s) found.")
    return 0
