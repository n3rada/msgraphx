# msgraphx/modules/graph/query.py
#
# Generic Microsoft Graph API query subcommand.
# Calls any Graph endpoint and returns the JSON response.
# Handles pagination automatically with --paginate.
#
# Required permissions: depends entirely on the endpoint called.
#
# Examples:
#   msgraphx --json graph /me
#   msgraphx --json graph /users --select id,displayName --filter "department eq 'IT'" --paginate
#   msgraphx --json graph /me/messages --top 10
#   msgraphx --json graph /me/sendMail --method POST --body '{"message": {...}}'

# Built-in imports
from __future__ import annotations

import argparse
import json
import sys

# External library imports
import httpx
from loguru import logger
from rich.json import JSON as RichJSON

# Local library imports
from ...core.context import GraphContext
from ...utils import output
from ...utils.console import console
from ...utils.errors import handle_graph_errors

_GRAPH_BASE = "https://graph.microsoft.com"


def _strip_odata(obj: dict) -> dict:
    return {k: v for k, v in obj.items() if not k.startswith("@odata.")}


def add_arguments(parser: "argparse.ArgumentParser") -> None:
    parser.add_argument(
        "path",
        type=str,
        help="Graph API path, e.g. /me or /users or /me/messages.",
    )

    parser.add_argument(
        "--method",
        "-X",
        type=str,
        default="GET",
        choices=["GET", "POST", "PATCH", "PUT", "DELETE"],
        metavar="METHOD",
        help="HTTP method (default: GET).",
    )

    parser.add_argument(
        "--filter",
        dest="odata_filter",
        type=str,
        default=None,
        metavar="EXPR",
        help="OData $filter expression, e.g. \"department eq 'IT'\".",
    )

    parser.add_argument(
        "--select",
        dest="select_fields",
        type=str,
        default=None,
        metavar="FIELDS",
        help="Comma-separated fields to return, e.g. id,displayName,mail.",
    )

    parser.add_argument(
        "--top",
        type=int,
        default=None,
        metavar="N",
        help="Max results per page (sets $top). Omit for server default.",
    )

    parser.add_argument(
        "--paginate",
        action="store_true",
        default=False,
        help="Follow @odata.nextLink to retrieve all pages.",
    )

    parser.add_argument(
        "--beta",
        action="store_true",
        default=False,
        help="Use the /beta endpoint instead of /v1.0.",
    )

    parser.add_argument(
        "--body",
        type=str,
        default=None,
        metavar="JSON",
        help="JSON body for POST/PATCH/PUT requests. Pass '-' to read from stdin.",
    )


@handle_graph_errors
async def run_with_arguments(context: "GraphContext", args: argparse.Namespace) -> int:
    token = await context.get_access_token()
    if not token:
        logger.error("No access token available for raw graph query.")
        return 1

    version = "beta" if getattr(args, "beta", False) else "v1.0"
    path = args.path if args.path.startswith("/") else f"/{args.path}"
    url = f"{_GRAPH_BASE}/{version}{path}"

    params: dict[str, str] = {}
    if args.odata_filter:
        params["$filter"] = args.odata_filter
    if args.select_fields:
        params["$select"] = args.select_fields
    if getattr(args, "top", None) is not None:
        params["$top"] = str(args.top)

    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json",
        "Content-Type": "application/json",
    }

    method = args.method.upper()
    body: dict | None = None
    if method in ("POST", "PATCH", "PUT") and args.body:
        raw = sys.stdin.read() if args.body == "-" else args.body
        try:
            body = json.loads(raw)
        except json.JSONDecodeError as exc:
            logger.error(f"Invalid JSON body: {exc}")
            return 1

    paginate = getattr(args, "paginate", False) or getattr(args, "fetch_all", False)

    collected: list[dict] = []
    raw_response: dict | None = None
    next_url: str | None = url
    first_page = True

    async with httpx.AsyncClient(verify=False) as client:
        while next_url:
            if method == "GET":
                resp = await client.get(
                    next_url,
                    headers=headers,
                    params=params if first_page else {},
                )
            elif method == "DELETE":
                resp = await client.delete(next_url, headers=headers)
            elif method == "POST":
                resp = await client.post(next_url, headers=headers, json=body)
            elif method == "PATCH":
                resp = await client.patch(next_url, headers=headers, json=body)
            elif method == "PUT":
                resp = await client.put(next_url, headers=headers, json=body)
            else:
                logger.error(f"Unsupported method: {method}")
                return 1

            first_page = False

            if resp.status_code == 204:
                # No content — success for DELETE/POST with no response body
                logger.success(f"{method} {path} — {resp.status_code} No Content")
                return 0

            if not resp.is_success:
                logger.error(f"Graph API error {resp.status_code}: {resp.text}")
                return 1

            try:
                data = resp.json()
            except Exception as exc:
                logger.error(f"Failed to parse Graph response: {exc}")
                return 1

            if "value" in data and isinstance(data["value"], list):
                page_items = data["value"]
                collected.extend(page_items)

                # Stream NDJSON items as they arrive
                if context.ndjson_output:
                    for item in page_items:
                        output.print_ndjson_item(_strip_odata(item))

                next_url = data.get("@odata.nextLink") if paginate else None
                logger.debug(
                    f"Page fetched: {len(page_items)} items, "
                    f"total so far: {len(collected)}"
                    + (f", next: {next_url}" if next_url else "")
                )
            else:
                raw_response = data
                next_url = None

    # Render results
    if raw_response is not None:
        if context.json_output:
            output.print_json(_strip_odata(raw_response))
        elif context.ndjson_output:
            output.print_ndjson_item(_strip_odata(raw_response))
        else:
            console.print(RichJSON(json.dumps(raw_response, default=str)))
        return 0

    # Collection response
    logger.success(f"{len(collected)} item(s) returned.")

    if context.json_output:
        output.print_json([_strip_odata(item) for item in collected])
    elif not context.ndjson_output:
        # ndjson items already streamed above; console gets raw
        console.print(RichJSON(json.dumps(collected, default=str)))

    return 0
