# msgraphx/modules/me/onenote.py
#
# List OneNote notebooks, sections, and pages for the current user.
# Pages often contain credentials, architecture diagrams, and internal docs.

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
from ...utils.html import strip_html
from ...utils.pagination import collect_all
from ...utils.roles import require_scopes

def add_arguments(parser: "argparse.ArgumentParser") -> None:
    mode = parser.add_mutually_exclusive_group()
    mode.add_argument(
        "--pages",
        action="store_true",
        help="List all pages across all notebooks.",
    )
    mode.add_argument(
        "--content",
        metavar="PAGE_ID",
        default=None,
        help="Fetch and display the HTML content of a specific page.",
    )

@handle_graph_errors
@require_scopes("Notes.Read")
async def run_with_arguments(
    context: "GraphContext", args: "argparse.Namespace"
) -> int:
    if context.is_app_only:
        logger.error("This module requires delegated authentication (user context).")
        return 1

    if args.content:
        return await _fetch_page_content(context, args.content)

    if args.pages:
        return await _list_pages(context)

    return await _list_notebooks(context)

async def _list_notebooks(context: "GraphContext") -> int:
    logger.info("Fetching OneNote notebooks")

    notebooks = await collect_all(context.graph_client.me.onenote.notebooks)

    if not notebooks:
        logger.info("No notebooks found.")
        if context.json_output:
            output.print_json([])
        return 0

    rows = []
    for nb in notebooks:
        modified = (
            nb.last_modified_date_time.strftime("%Y-%m-%d")
            if nb.last_modified_date_time else ""
        )
        rows.append({
            "id": nb.id,
            "display_name": nb.display_name,
            "is_shared": nb.is_shared,
            "modified": modified,
        })

    if context.json_output:
        output.print_json(rows)
        return 0

    if context.ndjson_output:
        for row in rows:
            output.print_ndjson_item(row)
        return 0

    table = Table(show_header=True, header_style="bold", box=None, padding=(0, 1))
    table.add_column("#", style="dim", justify="right", width=4)
    table.add_column("Notebook", min_width=40)
    table.add_column("Shared", style="dim", width=8)
    table.add_column("Modified", style="cyan", width=12)

    for i, row in enumerate(rows, 1):
        shared = "yes" if row["is_shared"] else ""
        table.add_row(str(i), row["display_name"] or "", shared, row["modified"])

    console.print("[bold]OneNote notebooks[/bold]")
    console.rule()
    console.print(table)
    logger.success(f"{len(rows)} notebook(s) found. Use --pages to list all pages.")
    return 0

async def _list_pages(context: "GraphContext") -> int:
    logger.info("Fetching all OneNote pages")

    pages = await collect_all(context.graph_client.me.onenote.pages)

    if not pages:
        logger.info("No pages found.")
        if context.json_output:
            output.print_json([])
        return 0

    rows = []
    for page in pages:
        modified = (
            page.last_modified_date_time.strftime("%Y-%m-%d")
            if page.last_modified_date_time else ""
        )
        notebook_name = ""
        section_name = ""
        if page.parent_section:
            section_name = page.parent_section.display_name or ""
        if page.parent_notebook:
            notebook_name = page.parent_notebook.display_name or ""

        rows.append({
            "id": page.id,
            "title": page.title,
            "notebook": notebook_name,
            "section": section_name,
            "modified": modified,
            "content_url": page.content_url,
        })

    cache.save_results(rows, key="onenote", identity=context.identity_hash)

    if context.json_output:
        output.print_json(rows)
        return 0

    if context.ndjson_output:
        for row in rows:
            output.print_ndjson_item(row)
        return 0

    table = Table(show_header=True, header_style="bold", box=None, padding=(0, 1))
    table.add_column("#", style="dim", justify="right", width=4)
    table.add_column("Title", min_width=40)
    table.add_column("Notebook / Section", style="dim", min_width=30)
    table.add_column("Modified", style="cyan", width=12)

    for i, row in enumerate(rows, 1):
        location = f"{row['notebook']} / {row['section']}" if row["notebook"] else row["section"]
        table.add_row(str(i), row["title"] or "(untitled)", location, row["modified"])

    console.print("[bold]OneNote pages[/bold]")
    console.rule()
    console.print(table)
    logger.success(f"{len(rows)} page(s) found. Use --content PAGE_ID to read a page.")
    return 0

async def _fetch_page_content(context: "GraphContext", page_id: str) -> int:
    logger.info(f"Fetching page content: {page_id}")

    content_bytes = await context.graph_client.me.onenote.pages.by_onenote_page_id(
        page_id
    ).content.get()

    if not content_bytes:
        logger.info("No content returned.")
        return 0

    html = content_bytes.decode("utf-8", errors="replace")
    text = strip_html(html)

    if context.json_output:
        output.print_json({"page_id": page_id, "html": html, "text": text})
        return 0

    if context.ndjson_output:
        output.print_ndjson_item({"page_id": page_id, "html": html, "text": text})
        return 0

    console.print("[bold]Page content[/bold]")
    console.rule()
    console.print(text)
    return 0
