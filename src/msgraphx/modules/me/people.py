# msgraphx/modules/me/people.py
#
# Return the most relevant people for the current user ranked by interaction
# frequency (email, meetings, shared docs). Faster than enumerating contacts
# for building a high-value target list for social engineering.
#
# Required delegated permissions:
#   People.Read

# Built-in imports
from __future__ import annotations

import argparse

# External library imports
from kiota_abstractions.base_request_configuration import RequestConfiguration
from loguru import logger
from msgraph.generated.users.item.people.people_request_builder import (
    PeopleRequestBuilder,
)
from rich.table import Table

# Local library imports
from ...core.context import GraphContext
from ...utils import output
from ...utils.console import console
from ...utils.errors import handle_graph_errors


def add_arguments(parser: "argparse.ArgumentParser") -> None:
    parser.add_argument(
        "--top",
        "-n",
        type=int,
        default=25,
        metavar="N",
        help="Number of people to return (default: 25).",
    )
    parser.add_argument(
        "--search",
        metavar="NAME",
        default=None,
        help="Filter by name (server-side $search).",
    )


async def fetch(
    context: GraphContext,
    top: int = 25,
    search: str | None = None,
) -> list[dict]:
    """Return the current user's people graph as plain dicts.

    Raises on API error. Callers are responsible for handling exceptions.
    """
    query_params = PeopleRequestBuilder.PeopleRequestBuilderGetQueryParameters(
        top=min(top, 1000),
        search=f'"{search}"' if search else None,
    )
    config = RequestConfiguration(query_parameters=query_params)

    result = await context.graph_client.me.people.get(request_configuration=config)
    people = (result.value or []) if result else []

    rows = []
    for person in people:
        email = ""
        phones_list = person.phones or []
        phone = phones_list[0].number if phones_list else ""

        scored_emails = person.scored_email_addresses or []
        if scored_emails:
            email = scored_emails[0].address or ""

        rows.append({
            "id": person.id,
            "display_name": person.display_name,
            "job_title": person.job_title,
            "department": person.department,
            "email": email,
            "phone": phone,
            "company": person.company_name,
            "office": person.office_location,
        })

    return rows


@handle_graph_errors
async def run_with_arguments(
    context: "GraphContext", args: "argparse.Namespace"
) -> int:
    if context.is_app_only:
        logger.error("This module requires delegated authentication (user context).")
        return 1

    logger.info("Fetching people graph")

    rows = await fetch(context, top=args.top, search=args.search)

    if not rows:
        logger.info("No people found.")
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
    table.add_column("Name", min_width=25)
    table.add_column("Email", style="cyan", min_width=30)
    table.add_column("Job title", style="dim", min_width=25)
    table.add_column("Department", style="dim", min_width=20)

    for i, row in enumerate(rows, 1):
        table.add_row(
            str(i),
            row["display_name"] or "",
            row["email"],
            row["job_title"] or "",
            row["department"] or "",
        )

    console.print("[bold]Your people graph[/bold]")
    console.rule()
    console.print(table)
    logger.success(f"{len(rows)} person/people found.")
    return 0
