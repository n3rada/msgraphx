# msgraphx/utils/tokens.py

# Built-in imports
import time
import json
import base64
from pathlib import Path
from datetime import datetime, timezone
import threading
import asyncio

# External library imports
from loguru import logger
import httpx
from azure.core.credentials import AccessToken


def parse_jwt(token: str) -> tuple[dict, dict, bytes]:
    """Decode JWT and return (header, body, signature)"""
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
    # Recognised source values: "file" | "env" | "arg"
    def __init__(
        self, access_token: str, refresh_token: str = None, source: str = "file"
    ):
        self._access_token = access_token
        self._refresh_token = refresh_token
        self._source = source

        try:
            self._header, self._payload, self._signature = parse_jwt(access_token)
        except Exception as e:
            logger.exception("Failed to parse JWT")
            raise

        self._expires_on = self._payload.get("exp")

        self._tenant_id = self._payload.get("iss")

        if self._tenant_id:
            self._tenant_id = self._tenant_id.strip("/").split("/").pop()
            logger.info(f"Tenant ID: {self._tenant_id}")

        exp_datetime = self.expiration_datetime
        human_date = exp_datetime.strftime("%A %d %b %Y, %H:%M:%S %Z")
        logger.info(f"JWT initialized, expires at {human_date}.")

    @property
    def audience(self) -> str:
        return self._payload.get("aud", "")

    @property
    def app_id(self) -> str:
        return self._payload.get("appid", "")

    @property
    def scope(self) -> str:
        return self._payload.get("scp", "")

    @property
    def access_token(self) -> str:
        return self._access_token

    @property
    def refresh_token(self) -> str:
        return self._refresh_token

    @property
    def payload(self) -> dict:
        return self._payload

    @property
    def expiration_datetime(self) -> datetime:
        return datetime.fromtimestamp(self._expires_on, tz=timezone.utc)

    @property
    def is_expired(self) -> bool:
        return datetime.now(timezone.utc).timestamp() > self._expires_on

    def expires_in(self) -> int:
        return max(0, int(self._expires_on - datetime.now(timezone.utc).timestamp()))

    def update_output_file(self) -> None:
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

    async def refresh_access_token(self, refresh_token: str):
        if not refresh_token:
            logger.warning(
                "No refresh token available - token refresh disabled. Re-authenticate when token expires."
            )
            return False

        response = httpx.post(
            url="https://login.microsoftonline.com/common/oauth2/v2.0/token",
            data={
                "client_id": "de8bc8b5-d9f9-48b1-a8ad-b748da725064",
                "scope": "openid https://graph.microsoft.com/.default offline_access",
                "grant_type": "refresh_token",
                "refresh_token": refresh_token,
            },
            headers={"Origin": "https://developer.microsoft.com"},
            verify=False,
        )

        if response.status_code != 200:
            logger.error(f"Failed to refresh token: {response.text}.")
            return False

        new_tokens = response.json()
        source = self._source  # preserve source across re-init

        self.__init__(
            access_token=new_tokens.get("access_token"),
            refresh_token=new_tokens.get("refresh_token"),
            source=source,
        )
        logger.success("Access token refreshed successfully.")
        return True

    def start_auto_refresh(self) -> threading.Thread | None:
        """Start a background daemon thread that silently refreshes the token.

        Refreshes 5 minutes before expiry. Returns the thread, or None if no
        refresh token is available.
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

    def get_token(self, *scopes, **kwargs) -> AccessToken:
        return AccessToken(self._access_token, self._expires_on)

    def __str__(self) -> str:
        return (
            f"Header:\n{json.dumps(self._header, indent=4, sort_keys=True)}\n\n"
            f"Payload:\n{json.dumps(self._payload, indent=4, sort_keys=True)}"
        )
