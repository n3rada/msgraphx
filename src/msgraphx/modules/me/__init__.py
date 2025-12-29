# Built-in imports
from typing import TYPE_CHECKING

# Third party library imports
from loguru import logger

# Local library imports
from . import groups
from msgraphx.utils.errors import handle_graph_errors


if TYPE_CHECKING:
    import argparse
    from msgraphx.core.context import GraphContext


def add_arguments(parser: "argparse.ArgumentParser"):
    subparsers = parser.add_subparsers(dest="subcommand", required=True)

    groups_parser = subparsers.add_parser(
        "groups", help="List current user's Microsoft 365 groups"
    )
    groups.add_arguments(groups_parser)


@handle_graph_errors
async def run_with_arguments(
    context: "GraphContext", args: "argparse.Namespace"
) -> int:
    if args.subcommand == "groups":
        return await groups.run_with_arguments(context, args)

    return 1
