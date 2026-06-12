# msgraphx/utils/errors.py

# Built-in imports
from __future__ import annotations

import re
import sys
import functools
import asyncio
from collections.abc import Callable
from typing import Any

# External library imports
from loguru import logger


class AuthenticationError(RuntimeError):
    """Raised when Graph API authentication fails (expired/invalid token)."""


class ForbiddenGraphError(RuntimeError):
    """Raised when the Graph API returns 403 Forbidden (insufficient permissions)."""

    def __init__(self, required: str | None, granted: str | None, raw_message: str):
        super().__init__(raw_message)
        self.required = required
        self.granted = granted
        self.raw_message = raw_message


def is_graph_auth_error(exc: Exception) -> bool:
    """Return True if the exception indicates an authentication failure."""
    try:
        # Import here to avoid import cycles at module import time
        from msgraph.generated.models.o_data_errors.o_data_error import ODataError
    except Exception:
        ODataError = None

    # Direct ODataError instance
    if ODataError is not None and isinstance(exc, ODataError):
        err_obj = exc
    else:
        # Try to find nested ODataError in common attributes
        err_obj = None
        for attr in ("__cause__", "__context__"):
            inner = getattr(exc, attr, None)
            if ODataError is not None and isinstance(inner, ODataError):
                err_obj = inner
                break

        if err_obj is None and hasattr(exc, "error"):
            err_obj = getattr(exc, "error")

    if not err_obj:
        return False

    # Extract code and message if available
    code = getattr(err_obj, "code", None)
    message = getattr(err_obj, "message", "")

    return code == "InvalidAuthenticationToken" or (
        isinstance(message, str) and "expired" in message.lower()
    )


def check_and_exit_for_auth_error(exc: Exception) -> None:
    """
    Inspect an exception for Microsoft Graph authentication errors (401/InvalidAuthenticationToken).
    If detected, log a clear message and exit the process with a non-zero code.

    This function is safe to call from generic exception handlers; it will exit
    the process only when the underlying error indicates an invalid/expired token.
    """
    try:
        # Import here to avoid import cycles at module import time
        from msgraph.generated.models.o_data_errors.o_data_error import ODataError
    except Exception:
        ODataError = None

    # Direct ODataError instance
    if ODataError is not None and isinstance(exc, ODataError):
        err_obj = exc
    else:
        # Try to find nested ODataError in common attributes
        err_obj = None
        for attr in ("__cause__", "__context__"):
            inner = getattr(exc, attr, None)
            if ODataError is not None and isinstance(inner, ODataError):
                err_obj = inner
                break

        if err_obj is None and hasattr(exc, "error"):
            err_obj = getattr(exc, "error")

    if not err_obj:
        return

    # Extract code and message if available
    code = getattr(err_obj, "code", None)
    message = getattr(err_obj, "message", "")

    if code == "InvalidAuthenticationToken" or (
        isinstance(message, str) and "expired" in message.lower()
    ):
        logger.error("Authentication failed: token invalid or expired. Stopping.")
        sys.exit(1)


def _odata_error_obj(exc: Exception):
    """Return the underlying ODataError-like object, or None."""
    try:
        from msgraph.generated.models.o_data_errors.o_data_error import ODataError
    except Exception:
        ODataError = None

    if ODataError is not None and isinstance(exc, ODataError):
        return exc
    for attr in ("__cause__", "__context__"):
        inner = getattr(exc, attr, None)
        if ODataError is not None and isinstance(inner, ODataError):
            return inner
    if hasattr(exc, "error"):
        return exc
    return None


def is_graph_forbidden_error(exc: Exception) -> bool:
    """Return True if exc represents a 403 Forbidden from the Graph API."""
    obj = _odata_error_obj(exc)
    if obj is None:
        return False
    err = getattr(obj, "error", None) or obj
    code = getattr(err, "code", None)
    if code == "Forbidden":
        return True
    # Some SDK versions surface the HTTP status directly
    status = getattr(exc, "response_status_code", None) or getattr(
        getattr(exc, "response", None), "status_code", None
    )
    return status == 403


def raise_if_forbidden(exc: Exception) -> None:
    """If exc is a 403 Forbidden, raise ForbiddenGraphError with parsed details."""
    if not is_graph_forbidden_error(exc):
        return

    obj = _odata_error_obj(exc)
    err = getattr(obj, "error", None) or obj
    raw = getattr(err, "message", None) or str(exc)

    required: str | None = None
    granted: str | None = None

    if raw:
        m = re.search(
            r"requires the following permissions:\s*(.+?)(?:\.\s*However|\Z)",
            raw,
            re.IGNORECASE | re.DOTALL,
        )
        if m:
            required = m.group(1).strip().rstrip(".")

        m = re.search(
            r"(?:permissions granted|following permissions granted):\s*(.+)",
            raw,
            re.IGNORECASE | re.DOTALL,
        )
        if m:
            granted = m.group(1).strip().rstrip(".")

    raise ForbiddenGraphError(required=required, granted=granted, raw_message=raw or "") from exc


def handle_graph_errors(func: Callable) -> Callable:
    """
    Decorator that catches exceptions from functions (sync or async), runs
    `check_and_exit_for_auth_error`, and re-raises the exception.

    Use on async functions as well; the decorator preserves coroutine behaviour.
    """

    if asyncio.iscoroutinefunction(func):

        @functools.wraps(func)
        async def async_wrapper(*args: Any, **kwargs: Any) -> Any:
            try:
                return await func(*args, **kwargs)
            except Exception as e:
                if is_graph_auth_error(e):
                    raise AuthenticationError("Graph authentication failed") from e
                raise

        return async_wrapper

    @functools.wraps(func)
    def sync_wrapper(*args: Any, **kwargs: Any) -> Any:
        try:
            return func(*args, **kwargs)
        except Exception as e:
            if is_graph_auth_error(e):
                raise AuthenticationError("Graph authentication failed") from e
            raise

    return sync_wrapper

