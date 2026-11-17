# tests/test_odata.py

from __future__ import annotations

import pytest

from msgraphx.utils import odata


@pytest.mark.parametrize(
    "value,expected",
    [
        ("alice@corp.com", "'alice@corp.com'"),
        ("o'brien@corp.com", "'o''brien@corp.com'"),
        ("x' or 1 eq 1 or mail eq 'y", "'x'' or 1 eq 1 or mail eq ''y'"),
        ("''", "''''''"),
        ("", "''"),
    ],
)
def test_quote_doubles_single_quotes(value: str, expected: str):
    assert odata.quote(value) == expected
