# Built-in imports
from typing import TYPE_CHECKING

# Third party library imports
from loguru import logger

# Local library imports
from . import search


if TYPE_CHECKING:
    import argparse
    from msgraphx.core.context import GraphContext


def add_arguments(parser: "argparse.ArgumentParser"):
    subparsers = parser.add_subparsers(dest="subcommand", required=True)

    search_parser = subparsers.add_parser(
        "search", help="Search Azure AD for groups, users, devices, and more"
    )
    search.add_arguments(search_parser)


async def run_with_arguments(
    context: "GraphContext", args: "argparse.Namespace"
) -> int:
    if args.subcommand == "search":
        return await search.run_with_arguments(context, args)

    return 1
