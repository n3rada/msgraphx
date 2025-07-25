# Built-in imports
import argparse

# Third party library imports
from loguru import logger
from msgraph import GraphServiceClient

# Local library imports
from graphx.modules.storage import search


def add_arguments(parser: argparse.ArgumentParser):
    subparsers = parser.add_subparsers(dest="subcommand", required=True)

    search_parser = subparsers.add_parser(
        "search", help="Search for files inside both SharePoint and OneDrive."
    )
    search.add_arguments(search_parser)

    # explorer_parser = subparsers.add_parser(
    #     "explore", help="Browse SharePoint/OneDrive"
    # )
    # explorer.add_arguments(explorer_parser)

    # download_parser = subparsers.add_parser(
    #     "download", help="Download files or folders"
    # )
    # download.add_arguments(download_parser)


async def run_with_arguments(
    graph_client: GraphServiceClient, args: argparse.Namespace
) -> int:
    if args.subcommand == "search":
        return await search.run_with_arguments(graph_client, args)
    # elif args.subcommand == "explore":
    #     return await explorer.run_with_arguments(args)
    # elif args.subcommand == "download":
    #     return await download.run_with_arguments(args)

    return 1
