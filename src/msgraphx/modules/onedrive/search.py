# msgraphx/modules/onedrive/search.py

# Built-in imports
from __future__ import annotations

import argparse

# External library imports
from msgraph.generated.models.search_content import SearchContent
from msgraph.generated.models.share_point_one_drive_options import SharePointOneDriveOptions

# Local library imports
from ...core.drive_search import DriveSearchBase
from ...core.context import GraphContext
from ...utils.errors import handle_graph_errors

HUNT_QUERIES: dict[str, str] = {
    "credentials": (
        "((filetype:key OR filetype:pem OR filetype:crt OR filetype:cer OR filetype:kdbx "
        "OR filetype:pfx) "
        "OR ((filetype:env OR filetype:cfg OR filetype:yaml OR filetype:yml OR filetype:secret) "
        'AND ("password" OR "passwd" OR "secret" OR "api_key" OR "access_key" OR "client_secret")))'
    ),
    "ssh": (
        "(filetype:pub OR filetype:pem OR filename:id_rsa OR filename:id_ecdsa OR "
        "filename:id_ed25519 OR filename:id_dsa OR filename:authorized_keys OR "
        'filename:known_hosts OR "BEGIN RSA PRIVATE KEY" OR '
        '"BEGIN OPENSSH PRIVATE KEY" OR "BEGIN EC PRIVATE KEY" OR "BEGIN PRIVATE KEY")'
    ),
    "office": (
        "filetype:doc OR filetype:docx OR filetype:xls OR filetype:xlsx OR "
        "filetype:ppt OR filetype:pptx OR filetype:pdf"
    ),
    "scripts": (
        "(filetype:ps1 OR filetype:sh OR filetype:bat OR filetype:cmd OR "
        "filetype:py OR filetype:rb OR filetype:pl OR filetype:ts)"
    ),
    "configs": (
        "((filetype:conf OR filetype:ini OR filetype:env OR filetype:yaml OR filetype:yml) "
        'AND ("password" OR "secret" OR "token" OR "credentials"))'
    ),
}


class OneDriveSearch(DriveSearchBase):
    """Search personal OneDrive drives (PrivateContent for app-only tokens)."""

    CACHE_KEY = "onedrive"
    LABEL = "OneDrive"
    HUNT_QUERIES = HUNT_QUERIES
    # For app-only: restrict to personal OneDrive drives only.
    # For delegated: this is ignored; the user sees their own OneDrive content.
    SCOPE = SharePointOneDriveOptions(include_content=SearchContent.PrivateContent)


async def fetch(
    context: GraphContext,
    query: str = "*",
    after: str | None = None,
    before: str | None = None,
) -> list[dict]:
    return await OneDriveSearch.fetch(context, query=query, after=after, before=before)


def add_arguments(parser: argparse.ArgumentParser) -> None:
    OneDriveSearch.add_arguments(parser)


@handle_graph_errors
async def run_with_arguments(
    context: GraphContext, args: argparse.Namespace
) -> int:
    return await OneDriveSearch.run_with_arguments(context, args)
