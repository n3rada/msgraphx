# msgraphx/utils/output.py

# Built-in imports
from __future__ import annotations

import json
import sys
from collections.abc import Mapping, Sequence
from typing import Any


def print_json(data: Sequence[Any] | Mapping[str, Any]) -> None:
    """Write data to stdout as indented JSON.

    Values that are not JSON-serializable are written with `str()`.

    Args:
        data: A list of records or a single record.
    """
    json.dump(data, sys.stdout, indent=2, default=str)
    sys.stdout.write("\n")
    sys.stdout.flush()


def print_ndjson_item(item: Mapping[str, Any]) -> None:
    """Write one record to stdout as a single JSON line (NDJSON).

    Args:
        item: The record to write. Non-serializable values use `str()`.
    """
    json.dump(item, sys.stdout, default=str)
    sys.stdout.write("\n")
    sys.stdout.flush()
