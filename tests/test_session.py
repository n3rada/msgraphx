# tests/test_session.py

from __future__ import annotations

import inspect
import types
from typing import Any

import pytest

from msgraphx import session as session_module
from msgraphx.core.context import GraphContext
from msgraphx.session import Session

_SENTINEL: list[dict[str, Any]] = [{"fake": True}]


def _validating_fake(original: Any) -> Any:
    """Return an async stand-in that rejects calls the original would reject."""
    signature = inspect.signature(original)

    async def fake(*args: Any, **kwargs: Any) -> list[dict[str, Any]]:
        signature.bind(*args, **kwargs)
        return _SENTINEL

    return fake


@pytest.fixture()
def fake_session(monkeypatch: pytest.MonkeyPatch) -> Session:
    """A Session whose module fetchers validate arguments instead of calling Graph."""
    for name, value in vars(session_module).items():
        if isinstance(value, types.ModuleType) and value.__name__.startswith(
            "msgraphx.modules."
        ):
            for attr, func in vars(value).items():
                if (
                    inspect.iscoroutinefunction(func)
                    and getattr(func, "__module__", None) == value.__name__
                    and not attr.startswith("_")
                ):
                    monkeypatch.setattr(value, attr, _validating_fake(func))
        elif inspect.iscoroutinefunction(value) and name.startswith("fetch_"):
            monkeypatch.setattr(session_module, name, _validating_fake(value))

    context = GraphContext(graph_client=None, is_app_only=False)  # type: ignore[arg-type]
    return Session(context)


def _wrapper_methods(fake_session: Session) -> list[tuple[str, Any]]:
    methods = []
    for ns_name in ("me", "aad", "outlook", "sharepoint", "teams"):
        namespace = getattr(fake_session, ns_name)
        for attr in dir(namespace):
            if attr.startswith("_"):
                continue
            method = getattr(namespace, attr)
            if inspect.iscoroutinefunction(method):
                methods.append((f"{ns_name}.{attr}", method))
    return methods


async def test_every_wrapper_calls_its_fetcher_with_valid_arguments(
    fake_session: Session,
):
    methods = _wrapper_methods(fake_session)
    assert len(methods) >= 25

    for label, method in methods:
        required = {
            name: "00000000-0000-0000-0000-000000000000"
            for name, param in inspect.signature(method).parameters.items()
            if param.default is inspect.Parameter.empty
        }
        assert await method(**required) is _SENTINEL, label


async def test_pim_roles_passes_state(fake_session: Session):
    assert await fake_session.aad.pim_roles(state="eligible") is _SENTINEL


@pytest.mark.parametrize(
    "kwargs", [{}, {"from_groups": True}, {"enrich": True}]
)
async def test_sites_modes(fake_session: Session, kwargs: dict[str, bool]):
    assert await fake_session.sharepoint.sites(**kwargs) is _SENTINEL


async def test_sites_rejects_conflicting_modes(fake_session: Session):
    with pytest.raises(ValueError):
        await fake_session.sharepoint.sites(from_groups=True, enrich=True)
