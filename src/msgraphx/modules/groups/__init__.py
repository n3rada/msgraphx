# Built-in imports
from __future__ import annotations

import argparse
from collections.abc import Sequence

# Local library imports
from . import list as list_cmd
from . import members, sites
from ...core.context import GraphContext
from ...utils.errors import handle_graph_errors


def add_arguments(
    parser: argparse.ArgumentParser,
    parents: Sequence[argparse.ArgumentParser] | None = None,
) -> None:
    """Register the `list`, `members`, and `sites` subcommands.

    Args:
        parser: Parser for the `groups` command.
        parents: Parent parsers whose shared options every subcommand
            inherits.
    """
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
    """Dispatch to the selected groups subcommand.

    Args:
        context: Shared Graph runtime context.
        args: Parsed arguments; `args.subcommand` selects the handler.

    Returns:
        The subcommand's exit code, or 1 for an unknown subcommand.

    Raises:
        AuthenticationError: The Graph token is invalid or expired.
    """
    sub = args.subcommand

    if sub in ("list", "ls"):
        return await list_cmd.run_with_arguments(context, args)
    if sub == "members":
        return await members.run_with_arguments(context, args)
    if sub == "sites":
        return await sites.run_with_arguments(context, args)

    return 1
