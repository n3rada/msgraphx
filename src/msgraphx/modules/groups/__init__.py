# Built-in imports
from __future__ import annotations

import argparse

# Local library imports
from . import list as list_cmd
from . import members, sites
from ...core.context import GraphContext
from ...utils.errors import handle_graph_errors


def add_arguments(
    parser: argparse.ArgumentParser, parents: list | None = None
) -> None:
    parents = parents or []
    subparsers = parser.add_subparsers(dest="subcommand", required=True)

    list_parser = subparsers.add_parser(
        "list",
        aliases=["ls"],
        parents=parents,
        help="List M365 Unified groups. Use --mine for your transitive membership.",
    )
    list_cmd.add_arguments(list_parser)

    members_parser = subparsers.add_parser(
        "members",
        parents=parents,
        help="List members of a group (direct or transitive).",
    )
    members.add_arguments(members_parser)

    sites_parser = subparsers.add_parser(
        "sites",
        parents=parents,
        help="List SharePoint sites owned by a group.",
    )
    sites.add_arguments(sites_parser)


@handle_graph_errors
async def run_with_arguments(context: GraphContext, args: argparse.Namespace) -> int:
    sub = args.subcommand

    if sub in ("list", "ls"):
        return await list_cmd.run_with_arguments(context, args)
    if sub == "members":
        return await members.run_with_arguments(context, args)
    if sub == "sites":
        return await sites.run_with_arguments(context, args)

    return 1
