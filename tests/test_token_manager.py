# tests/test_token_manager.py

from __future__ import annotations

import base64
import json
import time
from typing import Any

import httpx
import pytest

from msgraphx.utils import tokens
from msgraphx.utils.tokens import TokenManager


def _jwt(payload: dict[str, Any]) -> str:
    def b64url(data: bytes) -> str:
        return base64.urlsafe_b64encode(data).rstrip(b"=").decode()

    return ".".join(
        [
            b64url(json.dumps({"alg": "none"}).encode()),
            b64url(json.dumps(payload).encode()),
            b64url(b"sig"),
        ]
    )


def test_valid_token_exposes_claims_and_credential():
    exp = int(time.time()) + 3600
    manager = TokenManager(
        _jwt({"exp": exp, "aud": "https://graph.microsoft.com", "scp": "User.Read"})
    )

    assert not manager.is_expired
    assert manager.scope == "User.Read"
    assert manager.refresh_token is None
    access = manager.get_token("https://graph.microsoft.com/.default")
    assert access.expires_on == exp
    assert isinstance(access.expires_on, int)


def test_token_without_exp_is_rejected():
    with pytest.raises(ValueError, match="exp"):
        TokenManager(_jwt({"aud": "https://graph.microsoft.com"}))


@pytest.mark.parametrize("exp", ["1700000000", None, True])
def test_token_with_non_numeric_exp_is_rejected(exp: object):
    with pytest.raises(ValueError, match="exp"):
        TokenManager(_jwt({"exp": exp}))


def _patch_token_endpoint(
    monkeypatch: pytest.MonkeyPatch, status: int, body: dict[str, Any]
) -> list[httpx.Request]:
    seen: list[httpx.Request] = []

    def handler(request: httpx.Request) -> httpx.Response:
        seen.append(request)
        return httpx.Response(status, json=body)

    real_client = httpx.AsyncClient

    def fake_client(**kwargs: Any) -> httpx.AsyncClient:
        return real_client(transport=httpx.MockTransport(handler))

    monkeypatch.setattr(tokens.httpx, "AsyncClient", fake_client)
    return seen


async def test_refresh_adopts_new_tokens_and_keeps_source(monkeypatch: pytest.MonkeyPatch):
    old_exp = int(time.time()) + 60
    new_exp = int(time.time()) + 3600
    new_access = _jwt({"exp": new_exp, "scp": "Mail.Read"})
    seen = _patch_token_endpoint(
        monkeypatch, 200, {"access_token": new_access, "refresh_token": "rt-2"}
    )
    manager = TokenManager(_jwt({"exp": old_exp}), "rt-1", source="env")

    assert await manager.refresh_access_token(manager.refresh_token)

    assert b"refresh_token=rt-1" in seen[0].content
    assert manager.access_token == new_access
    assert manager.refresh_token == "rt-2"
    assert manager.scope == "Mail.Read"
    assert manager.expiration_datetime.timestamp() == new_exp
    assert manager._source == "env"


async def test_refresh_keeps_old_refresh_token_when_none_returned(
    monkeypatch: pytest.MonkeyPatch,
):
    new_access = _jwt({"exp": int(time.time()) + 3600})
    _patch_token_endpoint(monkeypatch, 200, {"access_token": new_access})
    manager = TokenManager(_jwt({"exp": int(time.time()) + 60}), "rt-1")

    assert await manager.refresh_access_token("rt-1")
    assert manager.refresh_token == "rt-1"


async def test_refresh_failure_keeps_current_tokens(monkeypatch: pytest.MonkeyPatch):
    _patch_token_endpoint(monkeypatch, 400, {"error": "invalid_grant"})
    access = _jwt({"exp": int(time.time()) + 60})
    manager = TokenManager(access, "rt-1")

    assert not await manager.refresh_access_token("rt-1")
    assert manager.access_token == access
    assert manager.refresh_token == "rt-1"
