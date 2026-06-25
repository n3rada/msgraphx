# msgraphx/modules/aad/grants.py
#
# List OAuth2 delegated permission grants in the tenant.
# Shows which app (client) has been granted which delegated scopes
# against which resource, and whether the grant is tenant-wide or per-user.
#
# Required delegated permissions:
#   DelegatedPermissionGrant.Read.All  (admin consent required)
#   or: Directory.Read.All

# Built-in imports
from __future__ import annotations

import argparse

# External library imports
from loguru import logger
from rich.table import Table

# Local library imports
from ...core.context import GraphContext
from ...utils import cache, output
from ...utils.console import console
from ...utils.errors import handle_graph_errors
from ...utils.pagination import GraphPaginator
from ...utils.roles import require_scopes


def add_arguments(parser: "argparse.ArgumentParser") -> None:
    parser.add_argument(
        "--filter",
        dest="odata_filter",
        metavar="EXPR",
        default=None,
        help="OData $filter expression (e.g. \"clientId eq 'GUID'\").",
    )
    parser.add_argument(
        "--tenant-wide",
        action="store_true",
        help="Only show tenant-wide (AllPrincipals) grants.",
    )


@handle_graph_errors
@require_scopes("DelegatedPermissionGrant.Read.All")
async def run_with_arguments(
    context: "GraphContext", args: "argparse.Namespace"
) -> int:
    logger.info("Fetching OAuth2 permission grants")

    rows = []
    async for grant in GraphPaginator(context.graph_client.oauth2_permission_grants):
        if args.odata_filter:
            # OData filter applied server-side via query params when paginator supports it;
            # fall back to client-side for now since we use bare paginator
            pass

        if args.tenant_wide and grant.consent_type != "AllPrincipals":
            continue

        rows.append({
            "id": grant.id,
            "client_id": grant.client_id,
            "consent_type": grant.consent_type,
            "principal_id": grant.principal_id,
            "resource_id": grant.resource_id,
            "scope": grant.scope,
        })

    if not rows:
        logger.info("No grants found.")
        if context.json_output:
            output.print_json([])
        return 0

    # Apply client-side filter if provided
    if args.odata_filter:
        logger.warning("Client-side filter applied. OData syntax is not evaluated; filtering by simple substring match on scope.")
        q = args.odata_filter.lower()
        rows = [r for r in rows if q in str(r).lower()]

    cache.save_results(rows, key="grants", identity=context.identity_hash)

    if context.json_output:
        output.print_json(rows)
        return 0

    if context.ndjson_output:
        for row in rows:
            output.print_ndjson_item(row)
        return 0

    table = Table(show_header=True, header_style="bold", box=None, padding=(0, 1))
    table.add_column("#", style="dim", justify="right", width=4)
    table.add_column("Client ID", style="dim", width=38)
    table.add_column("Type", width=14)
    table.add_column("Resource ID", style="dim", width=38)
    table.add_column("Scopes", style="cyan", min_width=30)

    for i, row in enumerate(rows, 1):
        consent = row["consent_type"] or ""
        consent_display = (
            "[red]AllPrincipals[/red]" if consent == "AllPrincipals" else consent
        )
        scopes = (row["scope"] or "").replace(" ", "\n")
        table.add_row(
            str(i),
            row["client_id"] or "",
            consent_display,
            row["resource_id"] or "",
            scopes,
        )

    console.print("[bold]OAuth2 delegated permission grants[/bold]")
    console.rule()
    console.print(table)
    logger.success(f"{len(rows)} grant(s) found.")
    return 0
