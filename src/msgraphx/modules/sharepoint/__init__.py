# Built-in imports
from __future__ import annotations

import argparse

# Local library imports
from . import search, download
from ...core.context import GraphContext
from ...utils.errors import handle_graph_errors


def add_arguments(parser: "argparse.ArgumentParser"):
    subparsers = parser.add_subparsers(dest="subcommand", required=True)

    search_parser = subparsers.add_parser(
        "search", help="Search for files inside both SharePoint and OneDrive."
    )
    search.add_arguments(search_parser)

    download_parser = subparsers.add_parser(
        "download", aliases=["dump"], help="Download all files from a drive"
    )
    download.add_arguments(download_parser)


@handle_graph_errors
async def run_with_arguments(
    context: "GraphContext", args: "argparse.Namespace"
) -> int:
    if args.subcommand == "search":
        return await search.run_with_arguments(context, args)
    elif args.subcommand in ("download", "dump"):
        return await download.run_with_arguments(context, args)

    return 1
