# msgraphx/modules/aad/roles.py
#
# List directory role assignments with principal and role names resolved.
# Maps every user/group/service-principal to every assigned directory role
# (Global Admin, Exchange Admin, Security Reader, etc.).
#
# Required delegated permissions:
#   RoleManagement.Read.Directory  (admin consent required)

# Built-in imports
from __future__ import annotations

import argparse

# External library imports
from kiota_abstractions.base_request_configuration import RequestConfiguration
from loguru import logger
from msgraph.generated.role_management.directory.role_assignments.role_assignments_request_builder import (
    RoleAssignmentsRequestBuilder,
)
from rich.table import Table

# Local library imports
from ...core.context import GraphContext
from ...utils import output
from ...utils.console import console
from ...utils.errors import handle_graph_errors
from ...utils.pagination import GraphPaginator


async def fetch(context: GraphContext, odata_filter: str | None = None) -> list[dict]:
    """Return directory role assignments as plain dicts.

    Raises on API error. Callers are responsible for handling exceptions.
    """
    query_params = RoleAssignmentsRequestBuilder.RoleAssignmentsRequestBuilderGetQueryParameters(
        expand=["roleDefinition", "principal"],
        filter=odata_filter or None,
        top=999,
    )
    config = RequestConfiguration(query_parameters=query_params)

    rows = []
    async for assignment in GraphPaginator(
        context.graph_client.role_management.directory.role_assignments, config
    ):
        role_name = ""
        if assignment.role_definition:
            role_name = assignment.role_definition.display_name or ""

        principal_name = ""
        principal_type = ""
        if assignment.principal:
            principal_name = getattr(assignment.principal, "display_name", "") or ""
            odata = getattr(assignment.principal, "odata_type", "") or ""
            principal_type = odata.split(".")[-1] if odata else ""

        scope = assignment.directory_scope_id or "/"

        rows.append({
            "assignment_id": assignment.id,
            "role_definition_id": assignment.role_definition_id,
            "role_name": role_name,
            "principal_id": assignment.principal_id,
            "principal_name": principal_name,
            "principal_type": principal_type,
            "scope": scope,
        })

    return rows


def add_arguments(parser: "argparse.ArgumentParser") -> None:
    parser.add_argument(
        "--filter",
        dest="odata_filter",
        metavar="EXPR",
        default=None,
        help="OData $filter expression (e.g. \"principalId eq 'GUID'\").",
    )


@handle_graph_errors
async def run_with_arguments(
    context: "GraphContext", args: "argparse.Namespace"
) -> int:
    logger.info("Fetching role assignments")

    rows = await fetch(context, odata_filter=args.odata_filter)

    if not rows:
        logger.info("No role assignments found.")
        if context.json_output:
            output.print_json([])
        return 0

    if context.json_output:
        output.print_json(rows)
        return 0

    if context.ndjson_output:
        for row in rows:
            output.print_ndjson_item(row)
        return 0

    table = Table(show_header=True, header_style="bold", box=None, padding=(0, 1))
    table.add_column("#", style="dim", justify="right", width=4)
    table.add_column("Role", min_width=30, style="cyan")
    table.add_column("Principal", min_width=35)
    table.add_column("Type", style="dim", width=16)
    table.add_column("Scope", style="dim", width=10)

    for i, row in enumerate(rows, 1):
        table.add_row(
            str(i),
            row["role_name"],
            row["principal_name"],
            row["principal_type"],
            row["scope"],
        )

    console.print("[bold]Directory role assignments[/bold]")
    console.rule()
    console.print(table)
    logger.success(f"{len(rows)} assignment(s) found.")
    return 0
