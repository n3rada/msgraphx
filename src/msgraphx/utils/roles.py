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
from functools import wraps

# External library imports
from loguru import logger

# Local library imports
from .tokens import parse_jwt

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
    try:
        _, payload, _ = parse_jwt(token)
        wids: list[str] = payload.get("wids", [])
        return [_ROLE_TEMPLATE_IDS[w] for w in wids if w in _ROLE_TEMPLATE_IDS]
    except Exception:
        return []


def _get_scp_scopes(token: str) -> frozenset[str]:
    try:
        _, payload, _ = parse_jwt(token)
        scp: str = payload.get("scp", "")
        return frozenset(scp.split()) if scp else frozenset()
    except Exception:
        return frozenset()


def require_roles(*roles: str):
    """Decorator: hard-fail if the delegated token lacks all of the listed directory roles.

    Passes if the token contains at least one of the named roles in its wids claim.
    App-only tokens are not checked (no wids claim).
    Stack outside @handle_graph_errors so auth errors still propagate normally.
    """
    def decorator(fn):
        @wraps(fn)
        async def wrapper(context, args):
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


def require_scopes(*scopes: str):
    """Decorator: warn if the delegated token is missing any of the listed OAuth scopes.

    Continues execution so the actual API 403 can surface with its own message.
    Skipped for app-only tokens (they use roles claim, not scp).
    """
    def decorator(fn):
        @wraps(fn)
        async def wrapper(context, args):
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
