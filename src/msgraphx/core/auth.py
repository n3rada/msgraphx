# msgraphx/core/auth.py
#
# Authenticated GraphContext factory. Lives here so both __init__.py
# and session.py can import it at module level without a circular dependency.

from __future__ import annotations

# External library imports
from azure.identity.aio import ClientSecretCredential
from msgraph.graph_service_client import GraphServiceClient

# Local library imports
from .context import GraphContext
from ..utils.tokens import TokenManager


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
    """Build an authenticated GraphContext for library use.

    App-only (client credentials) authentication is used when `tenant_id`
    and a client ID or secret are given; otherwise a delegated access token
    is required. A delegated-looking token that carries `roles` but no
    `scp` claim is treated as app-only. No Graph request is sent.

    Example:
        ctx = await create_context(access_token="eyJ...")
        ctx = await create_context(
            tenant_id="...", client_id="...", client_secret="..."
        )

    Args:
        access_token: Delegated JWT access token for Graph.
        refresh_token: Refresh token stored with the access token. It is not
            used for automatic renewal here.
        tenant_id: Tenant ID for app-only authentication.
        client_id: Application (client) ID for app-only authentication.
        client_secret: Client secret for app-only authentication.
        json_output: Value for `GraphContext.json_output`.
        ndjson_output: Value for `GraphContext.ndjson_output`.
        region: Search region for the Graph Search API.

    Returns:
        A context with a configured `GraphServiceClient` and token getter.
        `cached_user` and `raw_identity` are left unset.

    Raises:
        ValueError: If app-only credentials are incomplete, no access token
            is given for delegated auth, or the access token is malformed or
            expired.
    """
    if tenant_id and (client_id or client_secret):
        if not (client_id and client_secret):
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
