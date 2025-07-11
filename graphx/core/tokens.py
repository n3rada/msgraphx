# Built-in imports
import json
import base64
from pathlib import Path
from datetime import datetime, timezone
from typing import Callable, Optional

# External library imports
from loguru import logger


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
    def __init__(self, access_token: str, refresh_token: str = None):
        self._access_token = access_token
        self._refresh_token = refresh_token

        if self._refresh_token is not None:
            logger.info(
                "🔄 Refresh token provided, will be used for refreshing access token."
            )

        try:
            self._header, self._payload, self._signature = parse_jwt(access_token)
        except Exception as e:
            logger.error(f"❌ Failed to parse JWT: {e}")
            raise

        self._exp = self._payload.get("exp")

        self._tenant_id = self._payload.get("iss")
        print(self._tenant_id)
        if self._tenant_id:
            self._tenant_id = self._tenant_id.strip("/").split("/").pop()
            logger.info(f"🔗 Tenant ID: {self._tenant_id}")

        exp_datetime = self.expiration_datetime
        human_date = exp_datetime.strftime("%A %d %b %Y, %H:%M:%S %Z")
        logger.info(f"🔐 JWT initialized, expires at {human_date}.")

    @property
    def audience(self) -> str:
        return self._payload.get("aud", "")

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
        return datetime.fromtimestamp(self._exp, tz=timezone.utc)

    def is_expired(self) -> bool:
        return datetime.now(timezone.utc).timestamp() > self._exp

    def expires_in(self) -> int:
        return max(0, int(self._exp - datetime.now(timezone.utc).timestamp()))

    def update_output_file(self) -> None:
        output_file = Path(".roadtools_auth")
        output_file.unlink(missing_ok=True)

        data = {
            "tokenType": "Bearer",
            "tenantId": self._tenant_id,
            "expiresOn": self.expiration_datetime.strftime("%Y-%m-%d %H:%M:%S"),
            "_cliendId": "de8bc8b5-d9f9-48b1-a8ad-b748da725064",
            "accessToken": self._access_token,
            "refreshToken": self._refresh_token,
            "originheader": "https://developer.microsoft.com",
        }

        with output_file.open(mode="w", encoding="utf-8") as file_obj:
            json.dump(data, file_obj, indent=4)
            logger.success("💾 Saved refreshed tokens to .roadtools_auth")

    async def refresh_if_needed(
        self,
        threshold_seconds: int = 300,
        refresh_callback: Optional[Callable[[], str]] = None,
    ):
        if self.expires_in() <= threshold_seconds:
            if not refresh_callback:
                logger.warning(
                    "⚠️ Token is near expiration but no refresh callback provided."
                )
                return

            logger.info("♻️ Token is near expiration, refreshing...")
            try:
                result = await refresh_callback(self._refresh_token)
            except Exception as e:
                logger.exception(f"💥 Token refresh callback failed: {e}")
                raise RuntimeError("Token refresh failed unexpectedly.") from e

            if not result:
                logger.error(
                    "❌ Failed to refresh the token — refresh callback returned None."
                )
                raise RuntimeError("Token refresh failed: No token returned.")

            new_token, refresh_token = result
            self.__init__(new_token, refresh_token)

            self.update_output_file()
            logger.success("🔁 Token refreshed successfully.")

    def __str__(self) -> str:
        return (
            f"Header:\n{json.dumps(self._header, indent=4, sort_keys=True)}\n\n"
            f"Payload:\n{json.dumps(self._payload, indent=4, sort_keys=True)}"
        )
