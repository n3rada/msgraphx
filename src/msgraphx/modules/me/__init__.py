# Built-in imports
from __future__ import annotations

import argparse

# Local library imports
from . import calendar, drive, groups, onenote, people, planner, shared, trending, used
from ...core.context import GraphContext
from ...utils.errors import handle_graph_errors


def add_arguments(parser: "argparse.ArgumentParser"):
    subparsers = parser.add_subparsers(dest="subcommand", required=True)

    groups_parser = subparsers.add_parser(
        "groups", help="List current user's Microsoft 365 groups."
    )
    groups.add_arguments(groups_parser)

    trending_parser = subparsers.add_parser(
        "trending", help="Show documents trending around you."
    )
    trending.add_arguments(trending_parser)

    shared_parser = subparsers.add_parser(
        "shared", help="List documents shared with you (Insights API)."
    )
    shared.add_arguments(shared_parser)

    used_parser = subparsers.add_parser(
        "used", help="List documents you recently used (Insights API)."
    )
    used.add_arguments(used_parser)

    calendar_parser = subparsers.add_parser(
        "calendar", aliases=["cal"], help="List calendar events."
    )
    calendar.add_arguments(calendar_parser)

    onenote_parser = subparsers.add_parser(
        "onenote", aliases=["notes"], help="Browse OneNote notebooks and pages."
    )
    onenote.add_arguments(onenote_parser)

    planner_parser = subparsers.add_parser(
        "planner", aliases=["tasks"], help="List Planner tasks assigned to you."
    )
    planner.add_arguments(planner_parser)

    people_parser = subparsers.add_parser(
        "people", help="Show your people graph (top contacts by interaction)."
    )
    people.add_arguments(people_parser)

    drive_parser = subparsers.add_parser(
        "drive", aliases=["onedrive"], help="Browse and upload files in your personal OneDrive."
    )
    drive.add_arguments(drive_parser)


@handle_graph_errors
async def run_with_arguments(
    context: "GraphContext", args: "argparse.Namespace"
) -> int:
    sub = args.subcommand

    if sub == "groups":
        return await groups.run_with_arguments(context, args)
    if sub == "trending":
        return await trending.run_with_arguments(context, args)
    if sub == "shared":
        return await shared.run_with_arguments(context, args)
    if sub == "used":
        return await used.run_with_arguments(context, args)
    if sub in ("calendar", "cal"):
        return await calendar.run_with_arguments(context, args)
    if sub in ("onenote", "notes"):
        return await onenote.run_with_arguments(context, args)
    if sub in ("planner", "tasks"):
        return await planner.run_with_arguments(context, args)
    if sub == "people":
        return await people.run_with_arguments(context, args)
    if sub in ("drive", "onedrive"):
        return await drive.run_with_arguments(context, args)

    return 1
