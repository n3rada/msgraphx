# msgraphx/utils/errors.py

import sys
from loguru import logger
import functools
import asyncio
from typing import Callable, Any


class AuthenticationError(RuntimeError):
    """Raised when Graph API authentication fails (expired/invalid token)."""


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
        logger.error("🔒 Authentication failed: token invalid or expired. Stopping.")
        sys.exit(1)


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


# Backwards compatibility alias
handle_graph_errors = handle_graph_errors
