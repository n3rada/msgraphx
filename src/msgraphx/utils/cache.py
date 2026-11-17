# msgraphx/utils/cache.py

"""Lightweight result cache following XDG conventions.

Stores the most recent search results so users can download specific
items by index without re-querying the API.

Each search domain writes its own cache file:
  last_sharepoint.json  (sp search results)
  last_mail.json        (outlook search results)
"""

from __future__ import annotations

import json
import os
from collections.abc import Sequence
from pathlib import Path
from typing import Any

from loguru import logger

APP_NAME = "msgraphx"


def _get_data_dir(identity: str = "unknown") -> Path:
    """Return the cache directory for one operator identity.

    The base is `%APPDATA%` on Windows and `$XDG_DATA_HOME` (default
    `~/.local/share`) on POSIX. Partitioning by identity keeps several tokens
    on the same machine from overwriting each other's cached results. The
    directory is not created.

    Args:
        identity: The operator's `GraphContext.identity_hash`, a 16-character
            hex hash of the user OID (delegated) or `{tenant_id}:{client_id}`
            (app-only).

    Returns:
        The path `<base>/msgraphx/<identity>`.
    """
    if os.name == "nt":
        base = Path(os.environ.get("APPDATA") or Path.home())
    else:
        xdg = os.environ.get("XDG_DATA_HOME")
        base = Path(xdg) if xdg else Path.home() / ".local" / "share"

    return base / APP_NAME / identity


def save_results(
    items: Sequence[dict[str, Any]], key: str, identity: str = "unknown"
) -> None:
    """Replace the cached results for a search domain.

    Writes `last_<key>.json`, readable only by the owner on POSIX. A write
    failure is logged as a warning, not raised.

    Args:
        items: JSON-serializable result records.
        key: Search domain name, e.g. `sharepoint` or `mail`.
        identity: The operator's identity hash, see `_get_data_dir`.
    """
    data_dir = _get_data_dir(identity)
    data_dir.mkdir(parents=True, exist_ok=True)

    cache_file = data_dir / f"last_{key}.json"
    try:
        cache_file.write_text(json.dumps(items, indent=2), encoding="utf-8")
        if os.name != "nt":
            os.chmod(cache_file, 0o600)
        logger.debug(f"Cached {len(items)} result(s) to {cache_file}")
    except OSError as exc:
        logger.warning(f"Failed to cache results: {exc}")


def load_results(key: str, identity: str = "unknown") -> list[dict[str, Any]]:
    """Load the cached results for a search domain.

    Args:
        key: Search domain name, e.g. `sharepoint` or `mail`.
        identity: The operator's identity hash, see `_get_data_dir`.

    Returns:
        The cached records, or an empty list if there is no cache or it
        cannot be read or parsed.
    """
    cache_file = _get_data_dir(identity) / f"last_{key}.json"
    if not cache_file.is_file():
        return []

    try:
        return json.loads(cache_file.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        logger.warning(f"Failed to load cache: {exc}")
        return []


def parse_indices(spec: str, total: int) -> list[int]:
    """Convert a 1-based index spec such as `1,3-5,8` into 0-based indices.

    Args:
        spec: Comma-separated single indices (`3`) and inclusive ranges
            (`2-5`). Tokens that are not integers or ranges are skipped.
        total: Number of available items. Indices outside `1..total` are
            skipped.

    Returns:
        The valid indices, converted to 0-based, in spec order. Duplicates
        are kept.
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
