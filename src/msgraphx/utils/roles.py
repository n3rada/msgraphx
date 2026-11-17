# msgraphx/utils/roles.py
#
# Pre-flight permission checks via JWT claims.
#
# Two decorators for run_with_arguments:
#   @require_roles(*names)  — checks wids claim (directory roles).
#                             Hard-fail if none match. Delegated tokens only.
#   @require_scopes(*scopes) — checks scp claim (OAuth delegated scopes).
#                              Warning only; the API call may still succeed.
#                              Skipped for app-only tokens.
#
# wids contains role template IDs for roles the user currently holds,
# including PIM-activated assignments.

from __future__ import annotations

# Built-in imports
import argparse
from collections.abc import Awaitable, Callable
from functools import wraps
from typing import TypeAlias

# External library imports
from loguru import logger

# Local library imports
from ..core.context import GraphContext
from .tokens import parse_jwt

_ModuleEntryPoint: TypeAlias = Callable[
    [GraphContext, argparse.Namespace], Awaitable[int]
]

# Well-known Entra directory role template IDs.
# Source: https://learn.microsoft.com/en-us/entra/identity/role-based-access-control/permissions-reference
_ROLE_TEMPLATE_IDS: dict[str, str] = {
    "62e90394-69f5-4237-9190-012177145e10": "Global Administrator",
    "f2ef992c-3afb-46b9-b7cf-a126ee74c451": "Global Reader",
    "5d6b6bb7-de71-4623-b4af-96380a352509": "Security Reader",
    "194ae4cb-b126-40b2-bd5b-6091b380977d": "Security Administrator",
    "b1be1c3e-b65d-4f19-8427-f6fa0d97feb9": "Conditional Access Administrator",
    "9f06204d-73c1-4d4c-880a-6edb90606fd8": "Devices Administrator",
    "e8611ab8-c189-46e8-94e1-60213ab1f814": "Privileged Role Administrator",
    "7be44c8a-adaf-4e2a-84d6-ab2649e08a13": "Privileged Authentication Administrator",
    "29232cdf-9323-42fd-ade2-1d097af3e4de": "Exchange Administrator",
    "966707d0-3269-4727-9be2-8c3a10f19b9d": "Password Administrator",
    "c4e39bd9-1100-46d3-8c65-fb160da0071f": "Authentication Administrator",
}


def _get_wids_roles(token: str) -> list[str]:
    """Return the names of well-known directory roles in the `wids` claim.

    Role template IDs missing from `_ROLE_TEMPLATE_IDS` are ignored. A token
    that cannot be decoded yields an empty list.
    """
    try:
        _, payload, _ = parse_jwt(token)
        wids: list[str] = payload.get("wids", [])
        return [_ROLE_TEMPLATE_IDS[w] for w in wids if w in _ROLE_TEMPLATE_IDS]
    except Exception:
        return []


def _get_scp_scopes(token: str) -> frozenset[str]:
    """Return the delegated scopes in the `scp` claim, or an empty set."""
    try:
        _, payload, _ = parse_jwt(token)
        scp: str = payload.get("scp", "")
        return frozenset(scp.split()) if scp else frozenset()
    except Exception:
        return frozenset()


def require_roles(
    *roles: str,
) -> Callable[[_ModuleEntryPoint], _ModuleEntryPoint]:
    """Build a decorator that requires one of the given directory roles.

    The decorated `run_with_arguments(context, args)` entry point runs only
    if the delegated token's `wids` claim holds at least one of `roles`.
    Otherwise it logs the required and held roles and returns 1 without
    calling the entry point. App-only tokens have no `wids` claim and are
    not checked. Apply it outside `@handle_graph_errors` so authentication
    errors still propagate normally.

    Args:
        *roles: Directory role display names, e.g. `"Global Reader"`, as
            listed in `_ROLE_TEMPLATE_IDS`.

    Returns:
        A decorator for module entry points.
    """
    def decorator(fn: _ModuleEntryPoint) -> _ModuleEntryPoint:
        @wraps(fn)
        async def wrapper(
            context: GraphContext, args: argparse.Namespace
        ) -> int:
            if context.is_app_only:
                return await fn(context, args)
            token = await context.get_access_token()
            if token:
                user_roles = _get_wids_roles(token)
                if not any(r in user_roles for r in roles):
                    logger.error(
                        f"Insufficient directory roles. "
                        f"Required (any one of): {', '.join(roles)}. "
                        f"Token has: {', '.join(user_roles) or 'no elevated roles'}."
                    )
                    return 1
            return await fn(context, args)
        return wrapper
    return decorator


def require_scopes(
    *scopes: str,
) -> Callable[[_ModuleEntryPoint], _ModuleEntryPoint]:
    """Build a decorator that warns when delegated scopes are missing.

    The decorated `run_with_arguments(context, args)` entry point logs a
    warning listing any of `scopes` absent from the token's `scp` claim,
    then always runs, so the Graph API's own 403 can surface. App-only
    tokens carry roles rather than scopes and are not checked.

    Args:
        *scopes: OAuth delegated scope names, e.g. `"User.Read.All"`.

    Returns:
        A decorator for module entry points.
    """
    def decorator(fn: _ModuleEntryPoint) -> _ModuleEntryPoint:
        @wraps(fn)
        async def wrapper(
            context: GraphContext, args: argparse.Namespace
        ) -> int:
            if not context.is_app_only:
                token = await context.get_access_token()
                if token:
                    token_scopes = _get_scp_scopes(token)
                    missing = [s for s in scopes if s not in token_scopes]
                    if missing:
                        logger.warning(
                            f"Token may be missing required scope(s): {', '.join(missing)}."
                        )
            return await fn(context, args)
        return wrapper
    return decorator
