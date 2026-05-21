# Built-in imports
from __future__ import annotations

import argparse

# Local library imports
from . import search
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
        help="Search Azure AD for groups, users, devices, and more.",
    )
    search.add_arguments(search_parser)


@handle_graph_errors
async def run_with_arguments(
    context: "GraphContext", args: "argparse.Namespace"
) -> int:
    if args.subcommand == "search":
        return await search.run_with_arguments(context, args)

    return 1
