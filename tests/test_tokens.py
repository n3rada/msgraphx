"""Tests for msgraphx.utils.tokens module."""

from __future__ import annotations

import base64
import json
import time

import pytest

from msgraphx.utils.tokens import parse_jwt


def _make_jwt(header: dict, payload: dict, signature: bytes = b"sig") -> str:
    """Build a minimal JWT string for testing."""

    def b64url(data: bytes) -> str:
        return base64.urlsafe_b64encode(data).rstrip(b"=").decode()

    return ".".join(
        [
            b64url(json.dumps(header).encode()),
            b64url(json.dumps(payload).encode()),
            b64url(signature),
        ]
    )


class TestParseJwt:
    def test_basic_decode(self):
        header = {"alg": "RS256", "typ": "JWT"}
        payload = {"sub": "user123", "exp": int(time.time()) + 3600}
        token = _make_jwt(header, payload)

        h, p, s = parse_jwt(token)
        assert h == header
        assert p == payload
        assert isinstance(s, bytes)

    def test_empty_payload(self):
        header = {"alg": "none"}
        payload = {}
        token = _make_jwt(header, payload)

        h, p, _ = parse_jwt(token)
        assert h == header
        assert p == {}

    def test_invalid_format_no_dots(self):
        with pytest.raises(ValueError, match="Invalid JWT format"):
            parse_jwt("not-a-jwt")

    def test_invalid_format_two_parts(self):
        with pytest.raises(ValueError, match="Invalid JWT format"):
            parse_jwt("part1.part2")

    def test_payload_fields_preserved(self):
        payload = {
            "iss": "https://sts.windows.net/tenant-id/",
            "aud": "https://graph.microsoft.com",
            "exp": 1700000000,
            "scp": "Mail.Read User.Read",
        }
        token = _make_jwt({"alg": "RS256"}, payload)
        _, p, _ = parse_jwt(token)
        assert p["scp"] == "Mail.Read User.Read"
        assert "tenant-id" in p["iss"]
