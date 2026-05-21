# Built-in imports
from __future__ import annotations

import argparse

# Local library imports
from . import download, search
from ...core.context import GraphContext
from ...utils.errors import handle_graph_errors


def add_arguments(
    parser: "argparse.ArgumentParser", parents: "list | None" = None
) -> None:
    parents = parents or []
    subparsers = parser.add_subparsers(dest="subcommand", required=True)

    search_parser = subparsers.add_parser(
        "search",
        parents=parents,
        help="Search for files inside both SharePoint and OneDrive.",
    )
    search.add_arguments(search_parser)

    download_parser = subparsers.add_parser(
        "download",
        aliases=["dump"],
        parents=parents,
        help="Download all files from a drive.",
    )
    download.add_arguments(download_parser)


@handle_graph_errors
async def run_with_arguments(
    context: "GraphContext", args: "argparse.Namespace"
) -> int:
    subcommand = args.subcommand

    if subcommand == "search":
        return await search.run_with_arguments(context, args)
    if subcommand in ("download", "dump"):
        return await download.run_with_arguments(context, args)

    return 1
