# tests/test_graph_refresh.py

from __future__ import annotations

import argparse
from unittest.mock import MagicMock

import pytest

from msgraphx.modules.graph import refresh


@pytest.mark.skipif(refresh.sys.platform != "win32", reason="Windows-only behavior")
def test_is_alive_never_signals_on_windows(monkeypatch):
    kill = MagicMock()
    monkeypatch.setattr(refresh.os, "kill", kill)

    assert refresh._is_alive(1234) is False
    kill.assert_not_called()


def test_daemon_rejected_on_windows(monkeypatch):
    monkeypatch.setattr(refresh.sys, "platform", "win32")
    monkeypatch.setattr(refresh, "_load_tokens", MagicMock(return_value=("at", "rt", "env")))
    run_loop = MagicMock()
    monkeypatch.setattr(refresh, "_run_loop", run_loop)
    args = argparse.Namespace(stop=False, status=False, daemon=True, token_file=None, json=False)

    assert refresh.run_command(args) == 1
    run_loop.assert_not_called()
