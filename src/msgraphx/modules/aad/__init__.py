# Built-in imports
from __future__ import annotations

import argparse

# Local library imports
from . import authmethods, ca, grants, roles, search
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

    authmethods_parser = subparsers.add_parser(
        "authmethods",
        aliases=["mfa"],
        parents=parents,
        help="Enumerate MFA / authentication methods registered for a user.",
    )
    authmethods.add_arguments(authmethods_parser)

    ca_parser = subparsers.add_parser(
        "ca",
        aliases=["policies"],
        parents=parents,
        help="List conditional access policies.",
    )
    ca.add_arguments(ca_parser)

    roles_parser = subparsers.add_parser(
        "roles",
        parents=parents,
        help="List directory role assignments.",
    )
    roles.add_arguments(roles_parser)

    grants_parser = subparsers.add_parser(
        "grants",
        parents=parents,
        help="List OAuth2 delegated permission grants.",
    )
    grants.add_arguments(grants_parser)


@handle_graph_errors
async def run_with_arguments(
    context: "GraphContext", args: "argparse.Namespace"
) -> int:
    sub = args.subcommand

    if sub == "search":
        return await search.run_with_arguments(context, args)
    if sub in ("authmethods", "mfa"):
        return await authmethods.run_with_arguments(context, args)
    if sub in ("ca", "policies"):
        return await ca.run_with_arguments(context, args)
    if sub == "roles":
        return await roles.run_with_arguments(context, args)
    if sub == "grants":
        return await grants.run_with_arguments(context, args)

    return 1
