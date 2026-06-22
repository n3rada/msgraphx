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
from pathlib import Path

from loguru import logger

APP_NAME = "msgraphx"


def _get_data_dir(identity: str = "unknown") -> Path:
    """Return the XDG data directory for msgraphx, partitioned by operator identity.

    On POSIX: $XDG_DATA_HOME/msgraphx/{identity}/ (defaults to ~/.local/share/msgraphx/{identity}/)
    On Windows: %APPDATA%/msgraphx/{identity}/

    The identity is a 16-char hex hash derived from the operator's OID (delegated)
    or "{tenant_id}:{client_id}" (app-only), preventing multiple tokens on the
    same machine from overwriting each other's cached results.
    """
    if os.name == "nt":
        base = Path(os.environ.get("APPDATA") or Path.home())
    else:
        xdg = os.environ.get("XDG_DATA_HOME")
        base = Path(xdg) if xdg else Path.home() / ".local" / "share"

    return base / APP_NAME / identity


def save_results(items: list[dict], key: str, identity: str = "unknown") -> None:
    """Cache the latest results for *key* under the operator identity subdirectory."""
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


def load_results(key: str, identity: str = "unknown") -> list[dict]:
    """Load cached results for *key* from the operator identity subdirectory.

    Returns an empty list if no cache exists or the file is corrupt.
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
