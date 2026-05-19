# tests/test_cache.py

from __future__ import annotations

import json

import pytest

from msgraphx.utils.cache import (
    _get_data_dir,
    load_results,
    parse_indices,
    save_results,
)


# --- parse_indices ---


@pytest.mark.parametrize(
    "spec,total,expected",
    [
        ("1", 5, [0]),
        ("3", 5, [2]),
        ("1,3,5", 5, [0, 2, 4]),
        ("2-4", 5, [1, 2, 3]),
        ("1,3-5,8", 10, [0, 2, 3, 4, 7]),
        # Out-of-range silently skipped
        ("0", 5, []),
        ("6", 5, []),
        ("4-7", 5, [3, 4]),
        # Invalid tokens skipped
        ("abc", 5, []),
        ("1,abc,3", 5, [0, 2]),
        ("a-b", 5, []),
        # Edge cases
        ("1-1", 5, [0]),
        ("5", 5, [4]),
    ],
)
def test_parse_indices(spec: str, total: int, expected: list[int]):
    assert parse_indices(spec, total) == expected


# --- save/load round-trip ---


@pytest.fixture()
def xdg_data_home(tmp_path, monkeypatch):
    """Override XDG_DATA_HOME so cache goes to a temp directory."""
    monkeypatch.setenv("XDG_DATA_HOME", str(tmp_path))
    return tmp_path


def test_save_and_load(xdg_data_home):
    items = [
        {
            "drive_id": "d-001",
            "item_id": "i-001",
            "name": "report.pdf",
            "size": 1024,
            "web_url": "https://example.com/report.pdf",
        },
        {
            "drive_id": "d-002",
            "item_id": "i-002",
            "name": "notes.docx",
            "size": None,
            "web_url": None,
        },
    ]

    save_results(items, key="sharepoint")
    loaded = load_results(key="sharepoint")

    assert loaded == items


def test_load_empty_returns_empty_list(xdg_data_home):
    assert load_results(key="sharepoint") == []


def test_save_overwrites_previous(xdg_data_home):
    save_results([{"drive_id": "a", "item_id": "1", "name": "old.txt", "size": 0, "web_url": None}], key="sharepoint")
    save_results([{"drive_id": "b", "item_id": "2", "name": "new.txt", "size": 100, "web_url": None}], key="sharepoint")

    loaded = load_results(key="sharepoint")
    assert len(loaded) == 1
    assert loaded[0]["name"] == "new.txt"


def test_load_corrupt_file_returns_empty(xdg_data_home):
    cache_file = _get_data_dir() / "last_sharepoint.json"
    cache_file.parent.mkdir(parents=True, exist_ok=True)
    cache_file.write_text("not valid json {{{", encoding="utf-8")

    assert load_results(key="sharepoint") == []


def test_cache_file_permissions(xdg_data_home):
    import os
    import stat

    save_results([{"drive_id": "x", "item_id": "y", "name": "f.txt", "size": 0, "web_url": None}], key="sharepoint")

    cache_file = _get_data_dir() / "last_sharepoint.json"
    if os.name != "nt":
        mode = stat.S_IMODE(cache_file.stat().st_mode)
        assert mode == 0o600
