# Built-in imports
from typing import TYPE_CHECKING

# Third party library imports
from loguru import logger

# Local library imports
from . import search, download


if TYPE_CHECKING:
    import argparse
    from msgraph import GraphServiceClient


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


async def run_with_arguments(
    graph_client: "GraphServiceClient", args: "argparse.Namespace"
) -> int:
    if args.subcommand == "search":
        return await search.run_with_arguments(graph_client, args)
    elif args.subcommand in ("download", "dump"):
        return await download.run_with_arguments(graph_client, args)

    return 1
