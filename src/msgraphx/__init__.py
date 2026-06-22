# Built-in imports
from __future__ import annotations

import tomllib
from importlib.metadata import PackageNotFoundError, version
from pathlib import Path

try:
    __version__ = version("msgraphx")
except PackageNotFoundError:
    try:
        pyproject_path = Path(__file__).parent.parent.parent / "pyproject.toml"
        with open(pyproject_path, "rb") as f:
            __version__ = tomllib.load(f)["project"]["version"] + "-dev"
    except (FileNotFoundError, KeyError):
        __version__ = "unknown"


# ---------------------------------------------------------------------------
# Public library API
# ---------------------------------------------------------------------------

# Re-export GraphContext so callers can type-hint without digging into core.
from .core.context import GraphContext  # noqa: E402
from .session import Session  # noqa: E402

# External library imports
from azure.identity.aio import ClientSecretCredential  # noqa: E402
from msgraph.graph_service_client import GraphServiceClient  # noqa: E402

# Local library imports
from .utils.tokens import TokenManager  # noqa: E402


async def create_context(
    access_token: str | None = None,
    refresh_token: str | None = None,
    tenant_id: str | None = None,
    client_id: str | None = None,
    client_secret: str | None = None,
    json_output: bool = False,
    ndjson_output: bool = False,
    region: str = "EMEA",
) -> GraphContext:
    """Build an authenticated GraphContext for programmatic use.

    Delegated (user token):
        ctx = await create_context(access_token="eyJ...")

    App-only (client credentials):
        ctx = await create_context(
            tenant_id="...", client_id="...", client_secret="..."
        )

    For library integration use the Session API (no internal paths needed):

        from msgraphx import Session

        session = await Session.create(access_token="eyJ...")
        people  = await session.me.people(top=25)
        events  = await session.me.calendar(top=50, after="2025-01-01T00:00:00")
        users   = await session.aad.users(query="admin")
        files   = await session.sharepoint.search("filetype:xlsx credentials")
    """
    if tenant_id and (client_id or client_secret):
        # App-only: client credentials flow
        if not all((tenant_id, client_id, client_secret)):
            raise ValueError("App-only auth requires tenant_id, client_id, and client_secret.")

        credentials = ClientSecretCredential(
            tenant_id=tenant_id,
            client_id=client_id,
            client_secret=client_secret,
        )
        client = GraphServiceClient(
            credentials=credentials,
            scopes=["https://graph.microsoft.com/.default"],
        )

        async def _app_token_getter() -> str:
            t = await credentials.get_token("https://graph.microsoft.com/.default")
            return t.token

        return GraphContext(
            graph_client=client,
            is_app_only=True,
            region=region,
            json_output=json_output,
            ndjson_output=ndjson_output,
            token_getter=_app_token_getter,
        )

    # Delegated: bearer token
    if not access_token:
        raise ValueError("Delegated auth requires access_token.")

    token_mgr = TokenManager(access_token, refresh_token)

    if token_mgr.is_expired:
        raise ValueError("Access token is expired.")

    is_app_only = bool(token_mgr.payload.get("roles") and not token_mgr.payload.get("scp"))

    scopes: frozenset[str] = frozenset()
    if not is_app_only and token_mgr.scope:
        scopes = frozenset(token_mgr.scope.split())

    async def _delegated_token_getter() -> str:
        return token_mgr.access_token

    client = GraphServiceClient(token_mgr)

    return GraphContext(
        graph_client=client,
        is_app_only=is_app_only,
        region=region,
        token_scopes=scopes,
        json_output=json_output,
        ndjson_output=ndjson_output,
        token_getter=_delegated_token_getter,
    )


__all__ = ["__version__", "GraphContext", "Session", "create_context"]
