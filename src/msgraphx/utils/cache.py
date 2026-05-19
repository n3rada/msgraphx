# msgraphx/utils/cache.py

"""Lightweight result cache following XDG conventions.

Stores the most recent search results so users can download specific
items by index without re-querying the API.
"""

from __future__ import annotations

import json
import os
from pathlib import Path

from loguru import logger

APP_NAME = "msgraphx"


def _get_data_dir() -> Path:
    """Return the XDG data directory for msgraphx.

    On POSIX: $XDG_DATA_HOME/msgraphx (defaults to ~/.local/share/msgraphx)
    On Windows: %APPDATA%/msgraphx
    """
    if os.name == "nt":
        base = Path(os.environ.get("APPDATA") or Path.home())
    else:
        xdg = os.environ.get("XDG_DATA_HOME")
        base = Path(xdg) if xdg else Path.home() / ".local" / "share"

    return base / APP_NAME


def save_search_results(items: list[dict[str, str | int | None]]) -> None:
    """Cache the latest search results, overwriting any previous cache.

    Each item dict should contain at minimum:
        - drive_id: str
        - item_id: str
        - name: str
        - size: int | None
        - web_url: str | None
    """
    data_dir = _get_data_dir()
    data_dir.mkdir(parents=True, exist_ok=True)

    cache_file = data_dir / "last_search.json"
    try:
        cache_file.write_text(json.dumps(items, indent=2), encoding="utf-8")
        if os.name != "nt":
            os.chmod(cache_file, 0o600)
        logger.debug(f"Cached {len(items)} search result(s) to {cache_file}")
    except OSError as exc:
        logger.warning(f"Failed to cache search results: {exc}")


def load_search_results() -> list[dict[str, str | int | None]]:
    """Load the cached search results.

    Returns:
        List of cached item dicts, or empty list if no cache exists.
    """
    cache_file = _get_data_dir() / "last_search.json"
    if not cache_file.is_file():
        return []

    try:
        return json.loads(cache_file.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        logger.warning(f"Failed to load search cache: {exc}")
        return []


def parse_indices(spec: str, total: int) -> list[int]:
    """Parse a user index spec like '1,3,5-7' into 0-based indices.

    Supports:
        - Single: '3'
        - Comma-separated: '1,3,5'
        - Ranges: '2-5'
        - Mixed: '1,3-5,8'

    Out-of-range indices are silently skipped.
    """
    indices: list[int] = []
    for part in spec.split(","):
        part = part.strip()
        if "-" in part:
            bounds = part.split("-", 1)
            try:
                start = int(bounds[0])
                end = int(bounds[1])
                indices.extend(range(start, end + 1))
            except ValueError:
                continue
        else:
            try:
                indices.append(int(part))
            except ValueError:
                continue

    # Convert 1-based user input to 0-based, filter valid range
    return [i - 1 for i in indices if 1 <= i <= total]
