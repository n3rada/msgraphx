# msgraphx/utils/logbook.py

"""Logbook module for logging capabilities using Loguru.

Is this message about...

├─ Internal framework/library mechanics?
│  └─ Use TRACE
│
├─ Something the user might need to debug their usage?
│  └─ Use DEBUG
│
├─ Normal operational information?
│  └─ Use INFO
│
└─ Something went wrong or needs attention?
   └─ Use WARNING/ERROR
"""

# Built-in imports
from __future__ import annotations

import json
import os
import sys
import traceback
from datetime import timezone
from pathlib import Path
from typing import TYPE_CHECKING, Any

# Third party library imports
from loguru import logger

if TYPE_CHECKING:
    from loguru import Record


def _format_message(record: Record) -> str:
    """Build the colored stderr format template for a log record.

    Args:
        record: The Loguru record being formatted.

    Returns:
        A Loguru markup template with a UTC timestamp, a level symbol, and
        the message. Loguru fills in its `{time}`, `{message}`, and
        `{exception}` fields.
    """
    level_name = record["level"].name

    # Modern color palette (hex colors for better terminal support)
    trace_brown = "#8b7355"
    debug_blue = "#6c9bd1"
    info_white = "#ecf0f1"
    success_green = "#52c88a"
    warning_orange = "#f39c12"
    error_red = "#e74c3c"
    critical_magenta = "#c71585"
    time_gray = "#a5a5a5"

    # Map levels to symbols and colors
    symbols = {
        "TRACE": (f"<fg {trace_brown}>[*]</fg {trace_brown}>", trace_brown),
        "DEBUG": (f"<fg {debug_blue}>[•]</fg {debug_blue}>", debug_blue),
        "INFO": (f"<fg {info_white}>[i]</fg {info_white}>", info_white),
        "SUCCESS": (f"<fg {success_green}>[✓]</fg {success_green}>", success_green),
        "WARNING": (f"<fg {warning_orange}>[!]</fg {warning_orange}>", warning_orange),
        "ERROR": (f"<fg {error_red}>[✗]</fg {error_red}>", error_red),
        "CRITICAL": (
            f"<fg {critical_magenta}><bold>[⚠]</bold></fg {critical_magenta}>",
            critical_magenta,
        ),
    }

    symbol, color = symbols.get(level_name, ("[?]", "white"))

    # Professional format: full UTC timestamp + symbol + message
    # Note: We must return a string template, not format it yet
    return (
        f"<fg {time_gray}>{{time:YYYY-MM-DD HH:mm:ss.SSS!UTC}} (UTC)</fg {time_gray}> "
        f"{symbol} "
        f"<fg {color}>{{message}}</fg {color}>"
        "\n{exception}"
    )


_JSON_EXTRA_KEY = "_msgraphx_json"


def _serialize_record(record: Record) -> str:
    """Serialize a Loguru record into one compact JSON line.

    Args:
        record: The Loguru record to serialize. Extra keys starting with `_`
            are left out.

    Returns:
        A JSON object with `ts` (UTC, millisecond precision), `level`, `msg`,
        `module`, `function`, `line`, `exception` (formatted traceback or
        null), and `extra` when there is any.
    """
    exception = record["exception"]
    exception_text = None
    if exception is not None:
        exception_text = "".join(
            traceback.format_exception(
                exception.type, exception.value, exception.traceback
            )
        )

    payload: dict[str, Any] = {
        "ts": record["time"]
        .astimezone(timezone.utc)
        .isoformat(timespec="milliseconds")
        .replace("+00:00", "Z"),
        "level": record["level"].name,
        "msg": record["message"],
        "module": record["name"],
        "function": record["function"],
        "line": record["line"],
        "exception": exception_text,
    }

    extra = {
        key: value
        for key, value in record["extra"].items()
        if not key.startswith("_")
    }
    if extra:
        payload["extra"] = extra

    return json.dumps(
        payload, ensure_ascii=False, default=str, separators=(",", ":")
    )


def _format_json(record: Record) -> str:
    """Build the JSONL file-sink format template for a log record.

    Loguru treats the returned value as a `str.format` template, so the
    serialized JSON, which contains braces, is stored in the record's
    `extra` and the template only references it.

    Args:
        record: The Loguru record being formatted. Mutated: the serialized
            JSON is stored under `extra[_JSON_EXTRA_KEY]`.

    Returns:
        A template that renders as one JSON object followed by a newline.
    """
    record["extra"][_JSON_EXTRA_KEY] = _serialize_record(record)
    return "{extra[" + _JSON_EXTRA_KEY + "]}\n"


def _xdg_state_dir(app_name: str = "msgraphx") -> Path:
    """Return the platform-appropriate log directory.

    Uses `MSGRAPHX_LOG_DIR` when set, otherwise `%LOCALAPPDATA%` on Windows
    and `$XDG_STATE_HOME` (default `~/.local/state`) on POSIX. The directory
    is not created.

    Args:
        app_name: Subdirectory name under the platform base directory.

    Returns:
        The log directory path.
    """
    # Highest priority: explicit override
    override = os.getenv("MSGRAPHX_LOG_DIR")
    if override:
        return Path(override).expanduser().resolve()

    if os.name == "nt":
        base = Path(os.getenv("LOCALAPPDATA", str(Path.home() / "AppData" / "Local")))
        return (base / app_name / "logs").resolve()

    # POSIX: follow XDG
    xdg_state = os.getenv("XDG_STATE_HOME")
    if xdg_state:
        return Path(xdg_state).expanduser().resolve() / app_name / "logs"

    return Path.home() / ".local" / "state" / app_name / "logs"


def setup_logging(level: str = "INFO") -> None:
    """Replace all Loguru handlers with the msgraphx stderr and file sinks.

    Adds a colored, compact stderr sink and a rotating, zip-compressed
    `msgraphx.jsonl` file sink in the log directory, creating it if needed.
    Rotation size and retention come from `MSGRAPHX_LOG_MAX_BYTES` (default
    `10 MB`) and `MSGRAPHX_LOG_RETENTION_DAYS` (default 14).

    Args:
        level: Minimum level for both sinks: TRACE, DEBUG, INFO, SUCCESS,
            WARNING, ERROR, or CRITICAL, case-insensitive. Any other value
            falls back to INFO.
    """
    level = level.upper()

    # Validate log level
    valid_levels = ["TRACE", "DEBUG", "INFO", "SUCCESS", "WARNING", "ERROR", "CRITICAL"]
    if level not in valid_levels:
        level = "INFO"

    # Remove all Loguru handlers to avoid duplicates
    logger.remove()

    # Add custom formatted handler
    # enqueue=False for synchronous output to maintain ordering when using print()
    logger.add(
        sys.stderr,
        enqueue=False,
        backtrace=True,
        diagnose=True,
        level=level,
        format=_format_message,
        colorize=True,
    )

    # --- File handler (rotating JSONL, one object per line, UTC timestamps)
    log_dir = _xdg_state_dir()
    log_dir.mkdir(parents=True, exist_ok=True)
    log_file = log_dir / "msgraphx.jsonl"

    # Configurable rotation & retention
    max_bytes = os.getenv("MSGRAPHX_LOG_MAX_BYTES", "10 MB")
    retention_days = int(os.getenv("MSGRAPHX_LOG_RETENTION_DAYS", "14"))

    logger.add(
        log_file,
        format=_format_json,
        level=level,
        rotation=max_bytes,
        retention=f"{retention_days} days",
        compression="zip",
        encoding="utf-8",
        enqueue=True,  # Thread-safe
    )

    logger.trace(f"Logger initialized at level {level}")
    logger.trace(
        f"Log file: {log_file} (rotation {max_bytes}, retention {retention_days} days)"
    )
