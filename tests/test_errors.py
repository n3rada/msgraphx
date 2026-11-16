# tests/test_errors.py

from __future__ import annotations

import pytest
from msgraph.generated.models.o_data_errors.main_error import MainError
from msgraph.generated.models.o_data_errors.o_data_error import ODataError

from msgraphx.utils.errors import (
    AuthenticationError,
    ForbiddenGraphError,
    handle_graph_errors,
    is_graph_auth_error,
    raise_if_forbidden,
)


def _odata(code: str, message: str, status: int) -> ODataError:
    return ODataError(error=MainError(code=code, message=message), response_status_code=status)


def test_auth_error_detected_from_main_error_code():
    assert is_graph_auth_error(_odata("InvalidAuthenticationToken", "Access token validation failure.", 401))


def test_auth_error_detected_when_chained():
    try:
        try:
            raise _odata("InvalidAuthenticationToken", "Lifetime validation failed, the token is expired.", 401)
        except ODataError as inner:
            raise RuntimeError("wrapped") from inner
    except RuntimeError as exc:
        assert is_graph_auth_error(exc)


def test_non_auth_error_not_detected():
    assert not is_graph_auth_error(_odata("Request_ResourceNotFound", "Resource not found.", 404))


async def test_handle_graph_errors_translates_auth_error():
    @handle_graph_errors
    async def boom() -> None:
        raise _odata("InvalidAuthenticationToken", "Access token validation failure.", 401)

    with pytest.raises(AuthenticationError):
        await boom()


def test_raise_if_forbidden_parses_permissions():
    exc = _odata(
        "Forbidden",
        "This operation requires the following permissions: User.Read.All. "
        "However, the following permissions granted: User.Read.",
        403,
    )
    with pytest.raises(ForbiddenGraphError) as info:
        raise_if_forbidden(exc)
    assert info.value.required == "User.Read.All"
    assert info.value.granted == "User.Read"
