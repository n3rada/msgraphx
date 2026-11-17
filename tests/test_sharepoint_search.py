# tests/test_sharepoint_search.py

from __future__ import annotations

import argparse
from pathlib import Path
from unittest.mock import AsyncMock, MagicMock

from msgraph.generated.models.search_content import SearchContent

from msgraphx.core.context import GraphContext
from msgraphx.modules.sharepoint import search


def _context(is_app_only: bool = True) -> GraphContext:
    return GraphContext(graph_client=MagicMock(), is_app_only=is_app_only, region="EMEA")


def _args(**overrides: object) -> argparse.Namespace:
    values: dict[str, object] = {
        "query": "*",
        "filetype": None,
        "hunt": None,
        "before": None,
        "after": None,
        "save": None,
        "drive_id": None,
        "scope": None,
        "site": None,
        "my_groups": False,
        "visibility": None,
    }
    values.update(overrides)
    return argparse.Namespace(**values)


async def test_save_path_honors_site_and_scope(tmp_path: Path, monkeypatch):
    fetch = AsyncMock(return_value=[])
    monkeypatch.setattr(search.SharePointSearch, "fetch", fetch)
    monkeypatch.setattr(
        search, "_resolve_site", AsyncMock(return_value="https://t.sharepoint.com/sites/HR")
    )

    rc = await search.run_with_arguments(
        _context(), _args(save=str(tmp_path), site="HR", scope="onedrive")
    )

    assert rc == 0
    kwargs = fetch.await_args.kwargs
    assert kwargs["site_path"] == "https://t.sharepoint.com/sites/HR"
    assert kwargs["scope"].include_content == SearchContent.PrivateContent


async def test_display_and_save_paths_pass_same_scoping(tmp_path: Path, monkeypatch):
    fetch = AsyncMock(return_value=[])
    monkeypatch.setattr(search.SharePointSearch, "fetch", fetch)
    monkeypatch.setattr(
        search, "_resolve_site", AsyncMock(return_value="https://t.sharepoint.com/sites/HR")
    )

    await search.run_with_arguments(_context(), _args(site="HR", scope="sharepoint"))
    display = fetch.await_args.kwargs
    await search.run_with_arguments(
        _context(), _args(save=str(tmp_path), site="HR", scope="sharepoint")
    )
    saved = fetch.await_args.kwargs

    for key in ("site_path", "group_ids", "drive_id"):
        assert saved[key] == display[key], key
    # SDK models carry a per-instance backing store, so compare the field.
    assert saved["scope"].include_content == display["scope"].include_content


async def test_fetch_region_overrides_context_region(monkeypatch):
    fetch = AsyncMock(return_value=[])
    monkeypatch.setattr(search.SharePointSearch, "fetch", fetch)
    context = _context()

    await search.fetch(context, query="x", region="NAM")

    passed_context = fetch.await_args.args[0]
    assert passed_context.region == "NAM"
    assert context.region == "EMEA"


async def test_fetch_without_region_keeps_context(monkeypatch):
    fetch = AsyncMock(return_value=[])
    monkeypatch.setattr(search.SharePointSearch, "fetch", fetch)
    context = _context()

    await search.fetch(context, query="x")

    assert fetch.await_args.args[0] is context
