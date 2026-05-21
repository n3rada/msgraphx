# Built-in imports
from __future__ import annotations

import argparse

# Local library imports
from . import groups, trending
from ...core.context import GraphContext
from ...utils.errors import handle_graph_errors


def add_arguments(parser: "argparse.ArgumentParser"):
    subparsers = parser.add_subparsers(dest="subcommand", required=True)

    groups_parser = subparsers.add_parser(
        "groups", help="List current user's Microsoft 365 groups"
    )
    groups.add_arguments(groups_parser)

    trending_parser = subparsers.add_parser(
        "trending", help="Show documents trending around you"
    )
    trending.add_arguments(trending_parser)


@handle_graph_errors
async def run_with_arguments(
    context: "GraphContext", args: "argparse.Namespace"
) -> int:
    if args.subcommand == "groups":
        return await groups.run_with_arguments(context, args)
    if args.subcommand == "trending":
        return await trending.run_with_arguments(context, args)

    return 1
