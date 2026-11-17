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
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from loguru import logger

from ...utils.tokens import TokenManager


def _state_dir() -> Path:
    """Return the XDG state directory holding the daemon's PID and state."""
    xdg = os.environ.get("XDG_STATE_HOME")
    base = Path(xdg) if xdg else Path.home() / ".local" / "state"
    return base / "msgraphx"


def _pid_file() -> Path:
    return _state_dir() / "refresh.pid"


def _state_file() -> Path:
    return _state_dir() / "refresh.state.json"


def _write_state(expires_at: float, last_refreshed: float) -> None:
    """Persist the token expiry and last refresh time, ignoring write errors.

    Args:
        expires_at: Unix timestamp at which the current access token expires.
        last_refreshed: Unix timestamp of the last successful refresh, or
            0.0 if no refresh has happened yet.
    """
    try:
        _state_file().write_text(
            json.dumps({"expires_at": expires_at, "last_refreshed": last_refreshed})
        )
    except OSError:
        pass


def _read_state() -> dict[str, Any] | None:
    """Return the persisted daemon state, or None if missing or unreadable."""
    p = _state_file()
    if not p.exists():
        return None
    try:
        return json.loads(p.read_text())
    except (ValueError, OSError):
        return None


def _read_pid() -> int | None:
    """Return the daemon PID from the PID file, or None if missing or invalid."""
    p = _pid_file()
    if not p.exists():
        return None
    try:
        return int(p.read_text().strip())
    except (ValueError, OSError):
        return None


def _is_alive(pid: int) -> bool:
    """Return True if a process with the given PID exists.

    Always False on Windows: daemon mode is unavailable there, and
    `os.kill(pid, 0)` would terminate the process instead of probing it.
    """
    if sys.platform == "win32":
        return False
    try:
        os.kill(pid, 0)
        return True
    except (OSError, ProcessLookupError):
        return False


def _load_tokens(token_file: str | None) -> tuple[str | None, str | None, str]:
    """Resolve tokens from a roadtools token file, falling back to env vars.

    Args:
        token_file: Path to a `.roadtools_auth`-format JSON file. None uses
            `.roadtools_auth` in the current directory.

    Returns:
        A tuple (access_token, refresh_token, source), where either token may
        be None when not found and source is "file" or "env".
    """
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


async def _async_refresh_loop(access_token: str, refresh_token: str, source: str) -> None:
    """Refresh the token 5 minutes before each expiry until a refresh fails.

    Args:
        access_token: Current access token (JWT).
        refresh_token: Refresh token used to obtain new access tokens.
        source: Where the tokens came from ("file", "env", or "arg"); decides
            where refreshed tokens are written back.
    """
    token =TokenManager(access_token, refresh_token, source=source)
    _write_state(expires_at=time.time() + token.expires_in(), last_refreshed=0.0)
    logger.info("Refresh loop started.")
    while True:
        sleep_for = max(0, token.expires_in() - 300)
        if sleep_for > 0:
            logger.debug(f"Next refresh in {sleep_for}s.")
            await asyncio.sleep(sleep_for)
        ok = await token.refresh_access_token(token.refresh_token)
        if ok:
            token.update_output_file()
            _write_state(expires_at=time.time() + token.expires_in(), last_refreshed=time.time())
            logger.info("Token refreshed successfully.")
        else:
            logger.error("Token refresh failed. Stopping.")
            break


def _run_loop(access_token: str, refresh_token: str, source: str) -> None:
    """Run the refresh loop until a refresh fails, then remove PID and state.

    Args:
        access_token: Current access token (JWT).
        refresh_token: Refresh token used to obtain new access tokens.
        source: Where the tokens came from ("file", "env", or "arg").
    """
    # Clear any event loop inherited from the parent process after fork.
    asyncio.set_event_loop(None)
    try:
        asyncio.run(_async_refresh_loop(access_token, refresh_token, source))
    except Exception as exc:
        logger.error(f"Refresh loop error: {exc}")
    _pid_file().unlink(missing_ok=True)
    _state_file().unlink(missing_ok=True)


def _daemonize() -> None:
    """Detach from the terminal with a Unix double fork.

    The calling (parent) processes exit; only the grandchild returns, with
    stdin, stdout, and stderr redirected to the null device.

    Raises:
        NotImplementedError: On Windows, which has no `fork`/`setsid`.
    """
    if sys.platform == "win32":
        raise NotImplementedError("Daemon mode requires a POSIX system.")
    if os.fork() > 0:
        sys.exit(0)
    os.setsid()
    if os.fork() > 0:
        sys.exit(0)
    devnull = os.open(os.devnull, os.O_RDWR)
    for fd in (0, 1, 2):
        os.dup2(devnull, fd)
    os.close(devnull)


def is_running() -> bool:
    """Return True if a refresh daemon is active."""
    pid = _read_pid()
    return pid is not None and _is_alive(pid)


def add_arguments(parser: argparse.ArgumentParser) -> None:
    """Register the refresh command's arguments on a parser.

    Args:
        parser: Parser for the `refresh` command.
    """
    mode =parser.add_mutually_exclusive_group()
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
    parser.add_argument(
        "--json",
        action="store_true",
        default=False,
        help="Output --status as JSON.",
    )


def run_command(args: argparse.Namespace) -> int:
    """Run, daemonize, stop, or report the status of the token refresh loop.

    Without `--stop` or `--status`, this blocks until a refresh fails. With
    `--daemon`, the parent process exits and the detached child runs the loop.

    Args:
        args: Parsed `refresh` command arguments.

    Returns:
        0 on success or when there is nothing to stop, 1 when no access or
        refresh token is available or `--daemon` is used on Windows.
    """
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
        now = time.time()
        as_json = getattr(args, "json", False)

        alive = pid is not None and _is_alive(pid)
        if not alive:
            if pid is not None:
                _pid_file().unlink(missing_ok=True)
                _state_file().unlink(missing_ok=True)
            if as_json:
                print(json.dumps({"running": False}))
            else:
                print("not running")
            return 0

        state = _read_state() or {}
        last = state.get("last_refreshed", 0.0)
        expires_at = state.get("expires_at", 0.0)
        until = int(expires_at - now) if expires_at else None
        next_refresh = max(0, until - 300) if until is not None else None

        if as_json:
            print(json.dumps({
                "running": True,
                "pid": pid,
                "last_refreshed": datetime.fromtimestamp(last, tz=timezone.utc).isoformat() if last else None,
                "last_refreshed_ago_seconds": int(now - last) if last else None,
                "token_expires_at": datetime.fromtimestamp(expires_at, tz=timezone.utc).isoformat() if expires_at else None,
                "token_expires_in_seconds": until,
                "next_refresh_in_seconds": next_refresh,
            }))
        else:
            print(f"running (PID {pid})")
            if last:
                ago = int(now - last)
                print(f"  last refreshed : {datetime.fromtimestamp(last, tz=timezone.utc).strftime('%Y-%m-%d %H:%M:%S')} UTC  ({ago}s ago)")
            else:
                print("  last refreshed : not yet")
            if expires_at:
                exp_str = datetime.fromtimestamp(expires_at, tz=timezone.utc).strftime("%Y-%m-%d %H:%M:%S")
                print(f"  token expires  : {exp_str} UTC  (in {until}s)")
                print(f"  next refresh   : in ~{next_refresh}s")

        return 0

    access_token, refresh_token, source = _load_tokens(getattr(args, "token_file", None))
    if not access_token:
        logger.error("No access token found. Provide --token-file or set ACCESS_TOKEN.")
        return 1
    if not refresh_token:
        logger.error("No refresh token available. Cannot refresh without one.")
        return 1

    if args.daemon:
        if sys.platform == "win32":
            logger.error("--daemon is not supported on Windows. Run 'refresh' without --daemon instead.")
            return 1
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
