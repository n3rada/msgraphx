# msgraphx/modules/graph/refresh.py
#
# Background token refresh daemon.
#
#   msgraphx refresh              # run in foreground (blocking)
#   msgraphx refresh --daemon     # daemonize, write PID file, exit parent
#   msgraphx refresh --stop       # send SIGTERM to running daemon
#   msgraphx refresh --status     # show whether a daemon is running

from __future__ import annotations

import argparse
import asyncio
import json
import os
import signal
import sys
import time
from pathlib import Path

from loguru import logger

from ...utils.tokens import TokenManager


def _state_dir() -> Path:
    xdg = os.environ.get("XDG_STATE_HOME")
    base = Path(xdg) if xdg else Path.home() / ".local" / "state"
    return base / "msgraphx"


def _pid_file() -> Path:
    return _state_dir() / "refresh.pid"


def _read_pid() -> int | None:
    p = _pid_file()
    if not p.exists():
        return None
    try:
        return int(p.read_text().strip())
    except (ValueError, OSError):
        return None


def _is_alive(pid: int) -> bool:
    try:
        os.kill(pid, 0)
        return True
    except (OSError, ProcessLookupError):
        return False


def _load_tokens(token_file: str | None) -> tuple[str | None, str | None, str]:
    """Resolve tokens from a file, falling back to env vars."""
    path = Path(token_file) if token_file else Path(".roadtools_auth")
    if path.exists():
        try:
            data = json.loads(path.read_bytes())
            return data.get("accessToken"), data.get("refreshToken"), "file"
        except json.JSONDecodeError:
            logger.error(f"Failed to decode token file: {path}")

    env_token = os.environ.get("ACCESS_TOKEN")
    if env_token:
        return env_token, os.environ.get("REFRESH_TOKEN"), "env"

    return None, None, "file"


def _run_loop(access_token: str, refresh_token: str, source: str) -> None:
    """Refresh loop. Runs until a refresh fails, then removes the PID file."""
    loop = asyncio.new_event_loop()
    asyncio.set_event_loop(loop)
    try:
        token = TokenManager(access_token, refresh_token, source=source)
        while True:
            sleep_for = max(0, token.expires_in() - 300)
            if sleep_for > 0:
                logger.debug(f"Next refresh in {sleep_for}s.")
                time.sleep(sleep_for)
            ok = loop.run_until_complete(token.refresh_access_token(token.refresh_token))
            if ok:
                token.update_output_file()
            else:
                logger.error("Token refresh failed. Stopping.")
                break
    except Exception as exc:
        logger.error(f"Refresh loop error: {exc}")
    finally:
        loop.close()
    _pid_file().unlink(missing_ok=True)


def _daemonize() -> None:
    """Unix double-fork to fully detach from the terminal."""
    if os.fork() > 0:
        sys.exit(0)
    os.setsid()
    if os.fork() > 0:
        sys.exit(0)
    devnull = os.open(os.devnull, os.O_RDWR)
    for fd in (0, 1, 2):
        os.dup2(devnull, fd)
    os.close(devnull)


def add_arguments(parser: argparse.ArgumentParser) -> None:
    mode = parser.add_mutually_exclusive_group()
    mode.add_argument(
        "--daemon",
        action="store_true",
        default=False,
        help="Fork into the background and run the refresh loop as a daemon.",
    )
    mode.add_argument(
        "--stop",
        action="store_true",
        default=False,
        help="Stop a running refresh daemon.",
    )
    mode.add_argument(
        "--status",
        action="store_true",
        default=False,
        help="Show whether a refresh daemon is running.",
    )
    parser.add_argument(
        "--token-file",
        type=str,
        default=None,
        metavar="PATH",
        help="Token file to watch (default: .roadtools_auth in the current directory).",
    )


def run_command(args: argparse.Namespace) -> int:
    if args.stop:
        pid = _read_pid()
        if pid is None or not _is_alive(pid):
            if pid is not None:
                _pid_file().unlink(missing_ok=True)
            print("No refresh daemon running.")
            return 0
        os.kill(pid, signal.SIGTERM)
        print(f"Sent SIGTERM to PID {pid}.")
        return 0

    if args.status:
        pid = _read_pid()
        if pid is None:
            print("Refresh daemon: not running.")
            return 0
        if _is_alive(pid):
            print(f"Refresh daemon: running (PID {pid}).")
        else:
            _pid_file().unlink(missing_ok=True)
            print("Refresh daemon: not running (stale PID removed).")
        return 0

    access_token, refresh_token, source = _load_tokens(getattr(args, "token_file", None))
    if not access_token:
        logger.error("No access token found. Provide --token-file or set ACCESS_TOKEN.")
        return 1
    if not refresh_token:
        logger.error("No refresh token available. Cannot refresh without one.")
        return 1

    if args.daemon:
        pid = _read_pid()
        if pid and _is_alive(pid):
            print(f"Refresh daemon already running (PID {pid}).")
            return 0
        _daemonize()
        _state_dir().mkdir(parents=True, exist_ok=True)
        _pid_file().write_text(str(os.getpid()))
        _run_loop(access_token, refresh_token, source)
        return 0

    _run_loop(access_token, refresh_token, source)
    return 0
