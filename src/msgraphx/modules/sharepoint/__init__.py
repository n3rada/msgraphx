# Built-in imports
from __future__ import annotations

import argparse
from collections.abc import Sequence

# Local library imports
from . import download, search, sites
from ...core.context import GraphContext
from ...utils.errors import handle_graph_errors


def add_arguments(
    parser: argparse.ArgumentParser,
    parents: Sequence[argparse.ArgumentParser] | None = None,
) -> None:
    """Register the `search`, `sites`, and `download` subcommands.

    Args:
        parser: Parser for the `sharepoint` command.
        parents: Parent parsers whose shared options every subcommand
            inherits.
    """
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
    context: GraphContext, args: argparse.Namespace
) -> int:
    """Dispatch to the selected SharePoint subcommand.

    Args:
        context: Shared Graph runtime context.
        args: Parsed arguments; `args.subcommand` selects the handler.

    Returns:
        The subcommand's exit code, or 1 for an unknown subcommand.

    Raises:
        AuthenticationError: The Graph token is invalid or expired.
    """
    subcommand = args.subcommand

    if subcommand == "search":
        return await search.run_with_arguments(context, args)
    if subcommand == "sites":
        return await sites.run_with_arguments(context, args)
    if subcommand in ("download", "dump"):
        return await download.run_with_arguments(context, args)

    return 1
