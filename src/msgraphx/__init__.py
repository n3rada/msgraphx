# Built-in imports
from __future__ import annotations

import tomllib
from importlib.metadata import PackageNotFoundError, version
from pathlib import Path

try:
    __version__ = version("msgraphx")
except PackageNotFoundError:
    try:
        pyproject_path = Path(__file__).parent.parent.parent / "pyproject.toml"
        with open(pyproject_path, "rb") as f:
            __version__ = tomllib.load(f)["project"]["version"] + "-dev"
    except (FileNotFoundError, KeyError):
        __version__ = "unknown"


# ---------------------------------------------------------------------------
# Public library API
# ---------------------------------------------------------------------------

from .core.context import GraphContext
from .core.auth import create_context
from .session import Session

__all__ = ["__version__", "GraphContext", "Session", "create_context"]
