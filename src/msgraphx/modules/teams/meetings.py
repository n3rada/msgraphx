# msgraphx/modules/teams/meetings.py
#
# List online meetings and fetch transcripts for the current user.
# Transcripts are returned in VTT format (plain text with timestamps).

# Built-in imports
from __future__ import annotations

import argparse
from typing import Any

# External library imports
from loguru import logger
from rich.table import Table

# Local library imports
from ...core.context import GraphContext
from ...utils import output
from ...utils.console import console
from ...utils.errors import handle_graph_errors
from ...utils.pagination import GraphPaginator
from ...utils.roles import require_scopes

def add_arguments(parser: argparse.ArgumentParser) -> None:
    """Register the `teams meetings` options on a parser.

    Args:
        parser: Parser for the `teams meetings` subcommand.
    """
    parser.add_argument(
        "--top",
        "-n",
        type=int,
        default=25,
        metavar="N",
        help="Maximum number of meetings to list (default: 25).",
    )
    parser.add_argument(
        "--transcript",
        metavar="MEETING_ID",
        default=None,
        help="Fetch and print transcripts for a specific meeting ID.",
    )

async def fetch(context: GraphContext, top: int = 25) -> list[dict[str, Any]]:
    """List the signed-in user's online meetings.

    Args:
        context: Shared Graph runtime context (delegated).
        top: Maximum number of meetings to return.

    Returns:
        One dict per meeting with `id`, `subject`, `start`
        (YYYY-MM-DD HH:MM), `end` (HH:MM) and `organizer` (UPN, else display
        name).

    Raises:
        ODataError: A Graph request failed.
    """
    rows: list[dict[str, Any]] = []
    async for meeting in GraphPaginator(context.graph_client.me.online_meetings):
        start = ""
        end = ""
        if meeting.start_date_time:
            start = meeting.start_date_time.strftime("%Y-%m-%d %H:%M")
        if meeting.end_date_time:
            end = meeting.end_date_time.strftime("%H:%M")

        organizer = ""
        if meeting.participants and meeting.participants.organizer:
            org = meeting.participants.organizer
            if org.upn:
                organizer = org.upn
            elif org.identity and org.identity.user:
                organizer = org.identity.user.display_name or ""

        rows.append({
            "id": meeting.id,
            "subject": meeting.subject,
            "start": start,
            "end": end,
            "organizer": organizer,
        })

        if len(rows) >= top:
            break

    return rows

async def fetch_transcripts(
    context: GraphContext, meeting_id: str
) -> list[dict[str, Any]]:
    """Download every transcript of an online meeting.

    Makes one extra request per transcript to fetch its content.

    Args:
        context: Shared Graph runtime context (delegated).
        meeting_id: ID of the online meeting.

    Returns:
        One dict per transcript with `transcript_id`, `meeting_id`,
        `created` (YYYY-MM-DD HH:MM) and `content` (WebVTT text).

    Raises:
        ODataError: A Graph request failed.
    """
    result = await context.graph_client.me.online_meetings.by_online_meeting_id(
        meeting_id
    ).transcripts.get()

    transcripts = (result.value or []) if result else []

    items: list[dict[str, Any]] = []
    for transcript in transcripts:
        created = ""
        if transcript.created_date_time:
            created = transcript.created_date_time.strftime("%Y-%m-%d %H:%M")

        vtt_bytes = await (
            context.graph_client.me.online_meetings.by_online_meeting_id(meeting_id)
            .transcripts.by_call_transcript_id(transcript.id)
            .content.get()
        )

        vtt_text = ""
        if vtt_bytes:
            vtt_text = vtt_bytes.decode("utf-8", errors="replace")

        items.append({
            "transcript_id": transcript.id,
            "meeting_id": transcript.meeting_id,
            "created": created,
            "content": vtt_text,
        })

    return items

@handle_graph_errors
@require_scopes("OnlineMeetings.Read")
async def run_with_arguments(
    context: GraphContext, args: argparse.Namespace
) -> int:
    """List online meetings, or print one meeting's transcripts.

    Warns first if the token lacks `OnlineMeetings.Read`.

    Args:
        context: Shared Graph runtime context; must be delegated.
        args: Parsed arguments: `top`, and `transcript` (a meeting ID) to
            print transcripts instead of listing meetings.

    Returns:
        0 on success, 1 for app-only auth or if listing meetings fails.

    Raises:
        AuthenticationError: The Graph token is invalid or expired.
    """
    if context.is_app_only:
        logger.error("This module requires delegated authentication (user context).")
        return 1

    if args.transcript:
        return await _render_transcripts(context, args.transcript)

    return await _render_meetings(context, args.top)

async def _render_meetings(context: GraphContext, top: int) -> int:
    """Print online meetings as a table, JSON, or NDJSON.

    Args:
        context: Shared Graph runtime context (delegated).
        top: Maximum number of meetings to list.

    Returns:
        0 on success (including no meetings), 1 if the listing failed.
    """
    logger.info("Fetching online meetings")

    try:
        rows = await fetch(context, top=top)
    except Exception as exc:
        logger.error(f"Failed to list online meetings: {exc}")
        logger.info("Tip: OnlineMeetings.Read scope required, and meetings must have been created via Graph API or Teams.")
        return 1

    if not rows:
        logger.info("No meetings found.")
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
    table.add_column("Organizer", style="dim", min_width=25)
    table.add_column("ID", style="dim", min_width=30)

    for i, row in enumerate(rows, 1):
        table.add_row(
            str(i),
            row["subject"] or "(no subject)",
            row["start"],
            row["organizer"],
            row["id"] or "",
        )

    console.print("[bold]Online meetings[/bold]")
    console.rule()
    console.print(table)
    logger.success(f"{len(rows)} meeting(s) found. Use --transcript MEETING_ID to fetch a transcript.")
    return 0

async def _render_transcripts(context: GraphContext, meeting_id: str) -> int:
    """Print a meeting's transcripts as text, JSON, or NDJSON.

    Args:
        context: Shared Graph runtime context (delegated).
        meeting_id: ID of the online meeting.

    Returns:
        Always 0, including when the meeting has no transcripts.

    Raises:
        ODataError: A Graph request failed.
    """
    logger.info(f"Fetching transcripts for meeting: {meeting_id}")

    all_content = await fetch_transcripts(context, meeting_id)

    if not all_content:
        logger.info("No transcripts found for this meeting.")
        return 0

    if context.json_output:
        output.print_json(all_content)
        return 0

    if context.ndjson_output:
        for item in all_content:
            output.print_ndjson_item(item)
        return 0

    for item in all_content:
        console.print(f"[bold]Transcript {item['transcript_id']} ({item['created']})[/bold]")
        console.rule()
        console.print(item["content"])

    logger.success(f"{len(all_content)} transcript(s) fetched.")
    return 0
