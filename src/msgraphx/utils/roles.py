# msgraphx/utils/roles.py
#
# Entra directory role pre-flight checks via the wids JWT claim.
# wids lists role template IDs for roles actively assigned to the user
# (including PIM-activated ones). Only meaningful for delegated tokens;
# app-only tokens carry no wids and skip all checks.

from __future__ import annotations

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


def get_wids_roles(token: str) -> list[str]:
    """Return the list of directory role names present in the token's wids claim."""
    try:
        _, payload, _ = parse_jwt(token)
        wids: list[str] = payload.get("wids", [])
        return [_ROLE_TEMPLATE_IDS[w] for w in wids if w in _ROLE_TEMPLATE_IDS]
    except Exception:
        return []


async def require_any_role(context, required: list[str]) -> bool:
    """Return True if the token holder has at least one of the required directory roles.

    Returns True unconditionally for app-only tokens (no wids claim).
    Logs the gap and returns False when none of the required roles are present.
    """
    if context.is_app_only:
        return True

    token = await context.get_access_token()
    if not token:
        return True

    user_roles = get_wids_roles(token)
    if any(r in user_roles for r in required):
        return True

    logger.error(
        f"Insufficient directory roles. Required (any one of): {', '.join(required)}. "
        f"Token has: {', '.join(user_roles) or 'no elevated roles'}."
    )
    return False
