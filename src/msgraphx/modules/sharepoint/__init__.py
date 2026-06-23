# Built-in imports
from __future__ import annotations

import argparse

# Local library imports
from . import download, search, sites
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
        help="Search for files inside SharePoint team sites.",
    )
    search.add_arguments(search_parser)

    sites_parser = subparsers.add_parser(
        "sites",
        parents=parents,
        help="List SharePoint sites accessible via group membership.",
    )
    sites.add_arguments(sites_parser)

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
    if subcommand == "sites":
        return await sites.run_with_arguments(context, args)
    if subcommand in ("download", "dump"):
        return await download.run_with_arguments(context, args)

    return 1
