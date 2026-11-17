# msgraphx/utils/tokens.py

# Built-in imports
from __future__ import annotations

import time
import json
import base64
from pathlib import Path
from datetime import datetime, timezone
import threading
import asyncio
from typing import Any

# External library imports
from loguru import logger
import httpx
from azure.core.credentials import AccessToken


def parse_jwt(token: str) -> tuple[dict[str, Any], dict[str, Any], bytes]:
    """Decode a JWT without verifying its signature.

    Args:
        token: A compact-serialized JWT (`header.payload.signature`).

    Returns:
        A tuple (header, payload, signature), where header and payload are
        the decoded JSON objects and signature is the raw signature bytes.

    Raises:
        ValueError: If the token does not have three dot-separated parts or
            a part is not valid base64url-encoded JSON.
    """
    parts = token.split(".")
    if len(parts) != 3:
        raise ValueError("Invalid JWT format")

    def b64url_decode(data: str) -> bytes:
        padding = "=" * (-len(data) % 4)
        return base64.urlsafe_b64decode(data + padding)

    header = json.loads(b64url_decode(parts[0]))
    body = json.loads(b64url_decode(parts[1]))
    signature = b64url_decode(parts[2])
    return header, body, signature


class TokenManager:
    """A delegated Graph access token, with optional refresh-token renewal.

    Also serves as an Azure `TokenCredential` for `GraphServiceClient` via
    `get_token`.
    """

    # Recognised source values: "file" | "env" | "arg"
    def __init__(
        self,
        access_token: str,
        refresh_token: str | None = None,
        source: str = "file",
    ):
        """Decode the access token and log its tenant and expiry.

        Args:
            access_token: The JWT access token.
            refresh_token: Refresh token for renewal, or `None` to disable
                refreshing.
            source: Where the tokens came from, which decides where
                refreshed tokens are written back: `file`
                (`.roadtools_auth`), `env`, or `arg` (both environment
                variables).

        Raises:
            ValueError: If `access_token` is not a decodable JWT or has no
                numeric `exp` claim.
        """
        self._source = source
        self._load(access_token, refresh_token)

    def _load(self, access_token: str, refresh_token: str | None) -> None:
        """Adopt a token pair, decoding the access token's claims.

        Args:
            access_token: The JWT access token.
            refresh_token: Refresh token for renewal, or `None`.

        Raises:
            ValueError: If `access_token` is not a decodable JWT or has no
                numeric `exp` claim.
        """
        self._access_token = access_token
        self._refresh_token = refresh_token

        try:
            self._header, self._payload, self._signature = parse_jwt(access_token)
        except Exception:
            logger.exception("Failed to parse JWT")
            raise

        exp = self._payload.get("exp")
        if isinstance(exp, bool) or not isinstance(exp, (int, float)):
            raise ValueError("JWT has no numeric 'exp' claim.")
        self._expires_on: float = exp

        self._tenant_id = self._payload.get("iss")

        if self._tenant_id:
            self._tenant_id = self._tenant_id.strip("/").split("/").pop()
            logger.info(f"Tenant ID: {self._tenant_id}")

        exp_datetime = self.expiration_datetime
        human_date = exp_datetime.strftime("%A %d %b %Y, %H:%M:%S %Z")
        logger.info(f"JWT initialized, expires at {human_date}.")

    @property
    def audience(self) -> str:
        """The token's `aud` claim, or an empty string."""
        return self._payload.get("aud", "")

    @property
    def app_id(self) -> str:
        """The client application ID from the `appid` claim, or ''."""
        return self._payload.get("appid", "")

    @property
    def scope(self) -> str:
        """The space-separated delegated scopes (`scp` claim), or ''."""
        return self._payload.get("scp", "")

    @property
    def access_token(self) -> str:
        """The current access token."""
        return self._access_token

    @property
    def refresh_token(self) -> str | None:
        """The current refresh token, or `None` if refreshing is disabled."""
        return self._refresh_token

    @property
    def payload(self) -> dict[str, Any]:
        """The decoded access-token claims."""
        return self._payload

    @property
    def expiration_datetime(self) -> datetime:
        """The access token's expiry (`exp` claim) as an aware UTC datetime."""
        return datetime.fromtimestamp(self._expires_on, tz=timezone.utc)

    @property
    def is_expired(self) -> bool:
        """Whether the access token's expiry is in the past."""
        return datetime.now(timezone.utc).timestamp() > self._expires_on

    def expires_in(self) -> int:
        """Return the whole seconds until the access token expires, at least 0."""
        return max(0, int(self._expires_on - datetime.now(timezone.utc).timestamp()))

    def update_output_file(self) -> None:
        """Persist the current tokens back to where they were loaded from.

        For the `env` and `arg` sources this sets the `ACCESS_TOKEN` and
        `REFRESH_TOKEN` environment variables of the current process. For
        `file` it rewrites `.roadtools_auth` in the working directory. Does
        nothing when there is no refresh token.
        """
        if not self._refresh_token:
            logger.debug("Skipping token persistence - no refresh token available")
            return

        if self._source in ("env", "arg"):
            import os as _os

            _os.environ["ACCESS_TOKEN"] = self._access_token
            _os.environ["REFRESH_TOKEN"] = self._refresh_token
            logger.success(
                "Updated ACCESS_TOKEN / REFRESH_TOKEN env vars with refreshed tokens"
            )
            return

        # source == "file": write back to .roadtools_auth
        output_file = Path(".roadtools_auth")
        output_file.unlink(missing_ok=True)

        data = {
            "tokenType": "Bearer",
            "tenantId": self._tenant_id,
            "expiresOn": self.expiration_datetime.strftime("%Y-%m-%d %H:%M:%S"),
            "_clientId": "de8bc8b5-d9f9-48b1-a8ad-b748da725064",
            "accessToken": self._access_token,
            "refreshToken": self._refresh_token,
            "originheader": "https://developer.microsoft.com",
        }

        with output_file.open(mode="w", encoding="utf-8") as file_obj:
            json.dump(data, file_obj, indent=4)
            logger.success("Saved refreshed tokens to .roadtools_auth")

    async def refresh_access_token(self, refresh_token: str | None) -> bool:
        """Exchange a refresh token for new tokens and adopt them.

        Uses the Graph Explorer public client against the `common` endpoint.
        On success this manager adopts the new access token, and the new
        refresh token when one is returned (otherwise it keeps
        `refresh_token`). Tokens are not persisted; call
        `update_output_file` for that.

        Args:
            refresh_token: The refresh token to redeem, usually
                `self.refresh_token`.

        Returns:
            True if new tokens were obtained, False if there was no refresh
            token or the identity platform rejected the request.

        Raises:
            ValueError: If the returned access token cannot be decoded.
        """
        if not refresh_token:
            logger.warning(
                "No refresh token available - token refresh disabled. Re-authenticate when token expires."
            )
            return False

        async with httpx.AsyncClient(verify=False) as client:
            response = await client.post(
                url="https://login.microsoftonline.com/common/oauth2/v2.0/token",
                data={
                    "client_id": "de8bc8b5-d9f9-48b1-a8ad-b748da725064",
                    "scope": "openid https://graph.microsoft.com/.default offline_access",
                    "grant_type": "refresh_token",
                    "refresh_token": refresh_token,
                },
                headers={"Origin": "https://developer.microsoft.com"},
            )

        if response.status_code != 200:
            logger.error(f"Failed to refresh token: {response.text}.")
            return False

        new_tokens = response.json()
        self._load(
            access_token=new_tokens.get("access_token", ""),
            refresh_token=new_tokens.get("refresh_token") or refresh_token,
        )
        logger.success("Access token refreshed successfully.")
        return True

    def start_auto_refresh(self) -> threading.Thread | None:
        """Start a daemon thread that keeps the access token fresh.

        The thread refreshes 5 minutes before each expiry and persists the
        new tokens with `update_output_file`. It stops at the first failed
        refresh.

        Returns:
            The started thread, or `None` if there is no refresh token.
        """
        if not self._refresh_token:
            logger.debug("No refresh token. Auto-refresh disabled.")
            return None

        def _refresher() -> None:
            loop = asyncio.new_event_loop()
            asyncio.set_event_loop(loop)
            try:
                while True:
                    sleep_for = max(0, self.expires_in() - 300)
                    if sleep_for > 0:
                        logger.debug(f"Token refresh in {sleep_for}s.")
                        time.sleep(sleep_for)
                    try:
                        ok = loop.run_until_complete(
                            self.refresh_access_token(self._refresh_token)
                        )
                        if ok:
                            self.update_output_file()
                        else:
                            logger.error("Token refresh failed. Auto-refresh stopped.")
                            break
                    except Exception as exc:
                        logger.error(f"Token refresh error: {exc}")
                        break
            finally:
                loop.close()

        t = threading.Thread(target=_refresher, daemon=True, name="token-refresh")
        t.start()
        logger.info("Background token refresh started.")
        return t

    def get_token(self, *scopes: str, **kwargs: object) -> AccessToken:
        """Return the current access token as an Azure `AccessToken`.

        Implements `azure.core.credentials.TokenCredential` so the manager
        can be passed to `GraphServiceClient`. No token is requested: the
        requested scopes and options are ignored.

        Args:
            *scopes: Requested scopes; ignored.
            **kwargs: Credential options such as `claims`; ignored.

        Returns:
            The current access token and its expiry timestamp.
        """
        return AccessToken(self._access_token, int(self._expires_on))

    def __str__(self) -> str:
        """Return the decoded header and payload as indented JSON."""
        return (
            f"Header:\n{json.dumps(self._header, indent=4, sort_keys=True)}\n\n"
            f"Payload:\n{json.dumps(self._payload, indent=4, sort_keys=True)}"
        )
