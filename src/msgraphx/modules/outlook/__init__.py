# Built-in imports
from __future__ import annotations

import argparse

# Local library imports
from . import contacts, download, search, show
from ...core.context import GraphContext
from ...utils.errors import handle_graph_errors


def add_arguments(
    parser: "argparse.ArgumentParser", parents: "list | None" = None
) -> None:
    parents = parents or []
    subparsers = parser.add_subparsers(dest="subcommand", required=True)

    contacts_parser = subparsers.add_parser(
        "contacts", parents=parents, help="Browse Outlook contacts."
    )
    contacts.add_arguments(contacts_parser)

    download_parser = subparsers.add_parser(
        "download", parents=parents, help="Download email attachments."
    )
    download.add_arguments(download_parser)

    search_parser = subparsers.add_parser(
        "search", parents=parents, help="Search emails."
    )
    search.add_arguments(search_parser)

    show_parser = subparsers.add_parser(
        "show", parents=parents, help="Show a cached email."
    )
    show.add_arguments(show_parser)


@handle_graph_errors
async def run_with_arguments(
    context: "GraphContext", args: "argparse.Namespace"
) -> int:
    subcommand = args.subcommand

    if subcommand == "contacts":
        return await contacts.run_with_arguments(context, args)
    if subcommand == "download":
        return await download.run_with_arguments(context, args)
    if subcommand == "search":
        return await search.run_with_arguments(context, args)
    if subcommand == "show":
        return await show.run_with_arguments(context, args)

    return 1
