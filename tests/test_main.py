# tests/test_main.py

from __future__ import annotations

import argparse
import subprocess
import sys
from types import SimpleNamespace
from typing import Any

import pytest

from msgraphx import cli


def test_python_m_msgraphx_runs():
    result = subprocess.run(
        [sys.executable, "-m", "msgraphx", "--version"],
        capture_output=True,
        text=True,
        timeout=120,
    )

    assert result.returncode == 0, result.stderr
    assert "msgraphx" in result.stdout


async def test_main_returns_1_on_unexpected_exception(monkeypatch: pytest.MonkeyPatch):
    def boom() -> argparse.ArgumentParser:
        raise RuntimeError("boom")

    monkeypatch.setattr(cli, "build_parser", boom)

    assert await cli._main([]) == 1


async def test_authenticate_rejects_incomplete_client_credentials(
    monkeypatch: pytest.MonkeyPatch,
):
    for var in ("TENANT_ID", "CLIENT_ID", "CLIENT_SECRET", "ACCESS_TOKEN"):
        monkeypatch.delenv(var, raising=False)
    args = argparse.Namespace(
        access_token=None, tenant_id="tenant", client_id=None, client_secret=None
    )

    client, is_app_only, _, token_getter, err = await cli._authenticate(args)

    assert (client, is_app_only, token_getter, err) == (None, True, None, 1)


async def test_verify_connection_fails_when_me_is_empty():
    async def get() -> Any:
        return None

    graph_client: Any = SimpleNamespace(me=SimpleNamespace(get=get))

    assert await cli._verify_connection(graph_client, False, None, "EMEA") == (None, 1)
