# tests/test_sharepoint_sites.py

from __future__ import annotations

import argparse
from unittest.mock import AsyncMock, MagicMock

import pytest
from msgraph.generated.models.group import Group
from msgraph.generated.models.o_data_errors.main_error import MainError
from msgraph.generated.models.o_data_errors.o_data_error import ODataError
from msgraph.generated.models.site import Site

from msgraphx.core.context import GraphContext
from msgraphx.modules.sharepoint import sites
from msgraphx.utils.errors import AuthenticationError


def _odata(code: str, status: int) -> ODataError:
    return ODataError(error=MainError(code=code, message=code), response_status_code=status)


def _context(get: AsyncMock, is_app_only: bool = False) -> GraphContext:
    client = MagicMock()
    client.groups.by_group_id.return_value.sites.by_site_id.return_value.get = get
    return GraphContext(graph_client=client, is_app_only=is_app_only)


def _args(**overrides: object) -> argparse.Namespace:
    values: dict[str, object] = {"from_groups": False, "enrich": False}
    values.update(overrides)
    return argparse.Namespace(**values)


GROUP = Group(id="g-1", display_name="HR", visibility="Private")


async def test_group_site_returned():
    context = _context(AsyncMock(return_value=Site(id="s-1", name="HR")))

    pairs = await sites._resolve_group_sites(context, GROUP)

    assert pairs == [("s-1", {"group_id": "g-1", "group_name": "HR", "group_visibility": "Private"})]


async def test_group_without_site_returns_empty():
    context = _context(AsyncMock(side_effect=_odata("itemNotFound", 404)))

    assert await sites._fetch_group_sites(context, GROUP) == []


async def test_group_site_auth_error_propagates():
    context = _context(AsyncMock(side_effect=_odata("InvalidAuthenticationToken", 401)))

    with pytest.raises(ODataError):
        await sites._resolve_group_sites(context, GROUP)


async def test_group_site_unexpected_error_is_not_swallowed():
    context = _context(AsyncMock(side_effect=RuntimeError("network down")))

    with pytest.raises(RuntimeError):
        await sites._fetch_group_sites(context, GROUP)


async def test_enrich_uses_cache_without_search(monkeypatch):
    cached = [{"id": "s-1", "name": "HR", "web_url": "https://x"}]
    fetch = AsyncMock()
    monkeypatch.setattr(sites, "fetch", fetch)
    monkeypatch.setattr(sites.cache, "load_results", MagicMock(return_value=cached))
    save = MagicMock()
    monkeypatch.setattr(sites.cache, "save_results", save)
    enrich = AsyncMock(return_value=[{**cached[0], "access_via": "direct"}])
    monkeypatch.setattr(sites, "enrich_with_groups", enrich)
    monkeypatch.setattr(sites, "_print_sites", MagicMock())

    rc = await sites.run_with_arguments(_context(AsyncMock()), _args(enrich=True))

    assert rc == 0
    fetch.assert_not_awaited()
    enrich.assert_awaited_once()
    assert [c.kwargs["key"] for c in save.call_args_list] == ["sites_enriched"]


async def test_enrich_without_cache_searches_once(monkeypatch):
    found = [{"id": "s-1", "name": "HR", "web_url": "https://x"}]
    fetch = AsyncMock(return_value=found)
    monkeypatch.setattr(sites, "fetch", fetch)
    monkeypatch.setattr(sites.cache, "load_results", MagicMock(return_value=[]))
    save = MagicMock()
    monkeypatch.setattr(sites.cache, "save_results", save)
    monkeypatch.setattr(sites, "enrich_with_groups", AsyncMock(return_value=found))
    monkeypatch.setattr(sites, "_print_sites", MagicMock())

    rc = await sites.run_with_arguments(_context(AsyncMock()), _args(enrich=True))

    assert rc == 0
    fetch.assert_awaited_once()
    assert [c.kwargs["key"] for c in save.call_args_list] == ["sites", "sites_enriched"]


async def test_enrich_app_only_fails_before_search(monkeypatch):
    fetch = AsyncMock()
    monkeypatch.setattr(sites, "fetch", fetch)

    rc = await sites.run_with_arguments(
        _context(AsyncMock(), is_app_only=True), _args(enrich=True)
    )

    assert rc == 1
    fetch.assert_not_awaited()


async def test_run_converts_group_auth_error(monkeypatch):
    monkeypatch.setattr(sites, "get_user_m365_groups", AsyncMock(return_value=[GROUP]))
    context = _context(AsyncMock(side_effect=_odata("InvalidAuthenticationToken", 401)))

    with pytest.raises(AuthenticationError):
        await sites.run_with_arguments(context, _args(from_groups=True))
