# tests/test_aad_modules.py

from __future__ import annotations

import argparse
import json
from collections.abc import Sequence
from types import SimpleNamespace
from typing import Any

import pytest
from loguru import logger
from msgraph.generated.models.application import Application
from msgraph.generated.models.conditional_access_grant_control import (
    ConditionalAccessGrantControl,
)
from msgraph.generated.models.conditional_access_grant_controls import (
    ConditionalAccessGrantControls,
)
from msgraph.generated.models.conditional_access_policy import ConditionalAccessPolicy
from msgraph.generated.models.conditional_access_policy_state import (
    ConditionalAccessPolicyState,
)
from msgraph.generated.models.o_data_errors.main_error import MainError
from msgraph.generated.models.o_data_errors.o_data_error import ODataError

from msgraphx.modules.aad import apps, ca, grants, search
from msgraphx.utils import cache


class FakeCollection:
    """A single-page Graph collection builder that records each request."""

    def __init__(self, items: Sequence[Any] = (), error: Exception | None = None):
        self.items = list(items)
        self.error = error
        self.configs: list[Any] = []

    async def get(self, request_configuration: Any = None) -> Any:
        self.configs.append(request_configuration)
        if self.error is not None:
            raise self.error
        return SimpleNamespace(value=list(self.items), odata_next_link=None)

    def with_url(self, url: str) -> FakeCollection:
        return self


def _context(graph_client: Any, **overrides: Any) -> SimpleNamespace:
    defaults = {
        "graph_client": graph_client,
        "is_app_only": True,  # skips the token scope/role pre-flight checks
        "json_output": False,
        "ndjson_output": False,
        "identity_hash": "test",
    }
    return SimpleNamespace(**(defaults | overrides))


@pytest.fixture()
def warnings_log() -> list[str]:
    messages: list[str] = []
    sink_id = logger.add(lambda m: messages.append(m.record["message"]), level="WARNING")
    yield messages
    logger.remove(sink_id)


@pytest.fixture(autouse=True)
def no_cache_writes(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(cache, "save_results", lambda *a, **k: None)


# --- ca ---


def _policy(state: ConditionalAccessPolicyState, controls: list[ConditionalAccessGrantControl]) -> ConditionalAccessPolicy:
    return ConditionalAccessPolicy(
        id=state.value,
        display_name=state.value,
        state=state,
        grant_controls=ConditionalAccessGrantControls(built_in_controls=controls) if controls else None,
    )


@pytest.fixture()
def ca_context() -> SimpleNamespace:
    policies = FakeCollection([
        _policy(ConditionalAccessPolicyState.Enabled, [ConditionalAccessGrantControl.Mfa]),
        _policy(ConditionalAccessPolicyState.Disabled, [ConditionalAccessGrantControl.Block]),
        _policy(ConditionalAccessPolicyState.EnabledForReportingButNotEnforced, []),
    ])
    client = SimpleNamespace(policies=SimpleNamespace(conditional_access_policies=policies))
    return _context(client)


async def test_ca_fetch_uses_graph_wire_values(ca_context: SimpleNamespace):
    rows = await ca.fetch(ca_context)

    assert [r["state"] for r in rows] == [
        "enabled", "disabled", "enabledForReportingButNotEnforced",
    ]
    assert [r["grant_controls"] for r in rows] == [["mfa"], ["block"], []]


@pytest.mark.parametrize(
    "state,expected_id",
    [
        ("enabled", "enabled"),
        ("disabled", "disabled"),
        ("report", "enabledForReportingButNotEnforced"),
    ],
)
async def test_ca_fetch_filters_by_state(ca_context: SimpleNamespace, state: str, expected_id: str):
    rows = await ca.fetch(ca_context, state=state)
    assert [r["id"] for r in rows] == [expected_id]


# --- search ---


@pytest.mark.parametrize(
    "odata_filter,enabled_only,expected",
    [
        (None, False, None),
        ("department eq 'IT'", False, "department eq 'IT'"),
        (None, True, "accountEnabled eq true"),
        ("department eq 'IT'", True, "(department eq 'IT') and accountEnabled eq true"),
    ],
)
def test_search_build_filter(odata_filter: str | None, enabled_only: bool, expected: str | None):
    assert search._build_filter(odata_filter, enabled_only) == expected


async def test_search_users_enabled_only_filters_server_side():
    users = FakeCollection()
    await search.fetch_users(_context(SimpleNamespace(users=users)), "alice", enabled_only=True)

    params = users.configs[0].query_parameters
    assert params.filter == "accountEnabled eq true"
    assert params.search
    assert "accountEnabled" in params.select


async def test_search_users_custom_filter_selects_account_enabled():
    users = FakeCollection()
    await search.fetch_users(
        _context(SimpleNamespace(users=users)), "", odata_filter="department eq 'IT'",
    )

    params = users.configs[0].query_parameters
    assert params.filter == "department eq 'IT'"
    assert "accountEnabled" in params.select


@pytest.mark.parametrize("fetch,attr", [
    (search.fetch_devices, "devices"),
    (search.fetch_service_principals, "service_principals"),
])
async def test_search_enabled_only_applies_to_devices_and_sps(fetch: Any, attr: str):
    builder = FakeCollection()
    await fetch(_context(SimpleNamespace(**{attr: builder})), "srv", enabled_only=True)
    assert builder.configs[0].query_parameters.filter == "accountEnabled eq true"


async def test_search_enabled_only_warns_for_groups(warnings_log: list[str]):
    groups = FakeCollection()
    args = argparse.Namespace(
        hunt=None, query="admins", type="groups", contains=False, synced_only=False,
        enabled_only=True, odata_filter=None, save_dir=None,
    )

    assert await search.run_with_arguments(_context(SimpleNamespace(groups=groups)), args) == 0
    assert any("--enabled-only does not apply to groups" in m for m in warnings_log)
    assert groups.configs[0].query_parameters.filter is None


# --- grants ---


async def test_grants_filter_is_sent_to_graph(capsys: pytest.CaptureFixture[str]):
    grant = SimpleNamespace(
        id="g1", client_id="c1", consent_type="AllPrincipals",
        principal_id=None, resource_id="r1", scope="User.Read",
    )
    builder = FakeCollection([grant])
    context = _context(SimpleNamespace(oauth2_permission_grants=builder), json_output=True)
    args = argparse.Namespace(odata_filter="clientId eq 'c1'", tenant_wide=False)

    assert await grants.run_with_arguments(context, args) == 0
    assert builder.configs[0].query_parameters.filter == "clientId eq 'c1'"
    assert [r["id"] for r in json.loads(capsys.readouterr().out)] == ["g1"]


async def test_grants_without_filter_sends_no_query():
    builder = FakeCollection()
    context = _context(SimpleNamespace(oauth2_permission_grants=builder), json_output=True)

    await grants.run_with_arguments(context, argparse.Namespace(odata_filter=None, tenant_wide=False))
    assert builder.configs == [None]


# --- apps ---


async def test_apps_single_reports_forbidden_lookups_and_continues(
    capsys: pytest.CaptureFixture[str], warnings_log: list[str],
):
    app = Application(id="obj-1", app_id="client-1", display_name="Demo")
    forbidden = ODataError(
        error=MainError(code="Forbidden", message="Insufficient privileges."),
        response_status_code=403,
    )
    app_item = SimpleNamespace(get=_returning(app), owners=FakeCollection(error=forbidden))
    client = SimpleNamespace(
        applications=SimpleNamespace(by_application_id=lambda _id: app_item),
        service_principals=FakeCollection(),
    )
    context = _context(client, json_output=True)

    assert await apps.run_single(context, argparse.Namespace(app_id="obj-1")) == 0

    result = json.loads(capsys.readouterr().out)
    assert result["owners"] == []
    assert "servicePrincipal" not in result
    assert any("Could not fetch app owners" in m for m in warnings_log)


def _returning(value: Any) -> Any:
    async def get(*_args: Any, **_kwargs: Any) -> Any:
        return value
    return get
