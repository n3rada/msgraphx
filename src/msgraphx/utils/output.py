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
