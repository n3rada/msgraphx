# msgraphx/utils/output.py

# Built-in imports
from __future__ import annotations

import json
import sys


def print_json(data: list | dict) -> None:
    """Serialize data to JSON and write to stdout for piping."""
    json.dump(data, sys.stdout, indent=2, default=str)
    sys.stdout.write("\n")
    sys.stdout.flush()


def print_ndjson_item(item: dict) -> None:
    """Write one JSON object as a single line to stdout (NDJSON format)."""
    json.dump(item, sys.stdout, default=str)
    sys.stdout.write("\n")
    sys.stdout.flush()
