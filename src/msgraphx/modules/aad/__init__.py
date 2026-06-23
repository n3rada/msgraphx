# Built-in imports
from __future__ import annotations

import argparse

# Local library imports
from . import apps, authmethods, ca, enrich, grants, pim, roles, search
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
        help="Search Entra ID for groups, users, devices, and more.",
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

    user_parser = subparsers.add_parser(
        "user",
        parents=parents,
        help="Enriched user details: properties, owned objects, devices, app roles, OAuth2 grants.",
    )
    enrich.add_arguments_user(user_parser)

    group_parser = subparsers.add_parser(
        "group",
        parents=parents,
        help="Enriched group details: members, owners, drives, sites, app roles.",
    )
    enrich.add_arguments_group(group_parser)

    pim_parser = subparsers.add_parser(
        "pim",
        parents=parents,
        help="List PIM active and eligible role assignments.",
    )
    pim.add_arguments(pim_parser)

    app_parser = subparsers.add_parser(
        "app",
        parents=parents,
        help="Enriched app profile: credentials, permissions, service principal.",
    )
    apps.add_arguments_single(app_parser)

    apps_parser = subparsers.add_parser(
        "apps",
        parents=parents,
        help="List all app registrations with credential status.",
    )
    apps.add_arguments_bulk(apps_parser)


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
    if sub == "user":
        return await enrich.run_user(context, args)
    if sub == "group":
        return await enrich.run_group(context, args)
    if sub == "pim":
        return await pim.run_with_arguments(context, args)
    if sub == "app":
        return await apps.run_single(context, args)
    if sub == "apps":
        return await apps.run_bulk(context, args)

    return 1
