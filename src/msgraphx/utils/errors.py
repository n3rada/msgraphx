# msgraphx/utils/errors.py

# Built-in imports
from __future__ import annotations

import re
import sys
import functools
import inspect
from collections.abc import Callable
from typing import Any, TypeVar, cast

# External library imports
from loguru import logger
from msgraph.generated.models.o_data_errors.o_data_error import ODataError

_GraphCallable = TypeVar("_GraphCallable", bound=Callable[..., Any])


class AuthenticationError(RuntimeError):
    """A Graph API authentication failure: the token is invalid or expired."""


class ForbiddenGraphError(RuntimeError):
    """A 403 Forbidden from the Graph API: the token lacks a permission.

    Attributes:
        required: Permissions the API said it requires, if it said.
        granted: Permissions the API said the token has, if it said.
        raw_message: The API's error message.
    """

    def __init__(
        self, required: str | None, granted: str | None, raw_message: str
    ):
        """Initialize the error from the parsed 403 details.

        Args:
            required: Permissions the API requires, or `None` if unknown.
            granted: Permissions the token has, or `None` if unknown.
            raw_message: The API's error message, also used as the
                exception message.
        """
        super().__init__(raw_message)
        self.required = required
        self.granted = granted
        self.raw_message = raw_message


def is_graph_auth_error(exc: Exception) -> bool:
    """Check whether an exception is a Graph authentication failure.

    An `AuthenticationError` counts as one, so errors already translated by
    `handle_graph_errors` stay recognizable through nested calls. Otherwise
    looks at `exc` itself, then at the `ODataError` it was raised from or
    during, and treats an `InvalidAuthenticationToken` code or a message
    mentioning "expired" as an authentication failure.

    Args:
        exc: The exception to inspect.

    Returns:
        True if the token was rejected as invalid or expired.
    """
    if isinstance(exc, AuthenticationError):
        return True

    err_obj: object | None
    if isinstance(exc, ODataError):
        err_obj = exc
    else:
        err_obj = None
        for attr in ("__cause__", "__context__"):
            inner = getattr(exc, attr, None)
            if isinstance(inner, (ODataError, AuthenticationError)):
                err_obj = inner
                break

        if err_obj is None and hasattr(exc, "error"):
            err_obj = getattr(exc, "error")

    if isinstance(err_obj, AuthenticationError):
        return True
    if not err_obj:
        return False

    # ODataError carries code/message on its MainError payload (`.error`), not on itself
    err_obj = getattr(err_obj, "error", None) or err_obj
    code = getattr(err_obj, "code", None)
    message = getattr(err_obj, "message", "")

    return code == "InvalidAuthenticationToken" or (
        isinstance(message, str) and "expired" in message.lower()
    )


def check_and_exit_for_auth_error(exc: Exception) -> None:
    """Exit the process if an exception is a Graph authentication failure.

    Uses the same detection as `is_graph_auth_error`. Safe to call from
    generic exception handlers: other exceptions are ignored.

    Args:
        exc: The exception to inspect.

    Raises:
        SystemExit: With code 1, after logging an error, if the token was
            rejected as invalid or expired.
    """
    if is_graph_auth_error(exc):
        logger.error("Authentication failed: token invalid or expired. Stopping.")
        sys.exit(1)


def _odata_error_obj(exc: Exception) -> Exception | None:
    """Return the exception carrying the Graph error payload, if any.

    Args:
        exc: The exception to inspect.

    Returns:
        `exc` if it is an `ODataError` or has an `error` attribute, the
        `ODataError` it was raised from or during, or `None`.
    """
    if isinstance(exc, ODataError):
        return exc
    for attr in ("__cause__", "__context__"):
        inner = getattr(exc, attr, None)
        if isinstance(inner, ODataError):
            return inner
    if hasattr(exc, "error"):
        return exc
    return None


def is_graph_forbidden_error(exc: Exception) -> bool:
    """Check whether an exception is a 403 Forbidden from the Graph API.

    Args:
        exc: The exception to inspect.

    Returns:
        True if the Graph error code is `Forbidden` or the HTTP status of
        `exc` is 403.
    """
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
    """Re-raise a Graph 403 Forbidden as a `ForbiddenGraphError`.

    Does nothing if `exc` is not a 403, so callers can call it first in an
    `except` block and then handle or re-raise other errors.

    Args:
        exc: The caught exception.

    Raises:
        ForbiddenGraphError: If `exc` is a 403. The required and granted
            permissions are parsed from the API message when present, and
            `exc` is chained as the cause.
    """
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


def handle_graph_errors(func: _GraphCallable) -> _GraphCallable:
    """Decorate a function so Graph authentication failures become typed errors.

    Works on sync and async functions; an async function stays a coroutine
    function. Exceptions that `is_graph_auth_error` recognizes are re-raised
    as `AuthenticationError`, with the original chained as the cause; an
    `AuthenticationError` from a nested decorated call propagates as is. All
    other exceptions propagate unchanged.

    Args:
        func: The function to wrap.

    Returns:
        The wrapped function, with the same signature.
    """
    if inspect.iscoroutinefunction(func):

        @functools.wraps(func)
        async def async_wrapper(*args: Any, **kwargs: Any) -> Any:
            try:
                return await func(*args, **kwargs)
            except AuthenticationError:
                raise
            except Exception as e:
                if is_graph_auth_error(e):
                    raise AuthenticationError("Graph authentication failed") from e
                raise

        return cast(_GraphCallable, async_wrapper)

    @functools.wraps(func)
    def sync_wrapper(*args: Any, **kwargs: Any) -> Any:
        try:
            return func(*args, **kwargs)
        except AuthenticationError:
            raise
        except Exception as e:
            if is_graph_auth_error(e):
                raise AuthenticationError("Graph authentication failed") from e
            raise

    return cast(_GraphCallable, sync_wrapper)

