# msgraphx/core/context.py

from __future__ import annotations

import hashlib
from collections.abc import Awaitable, Callable
from dataclasses import dataclass, field

from msgraph.graph_service_client import GraphServiceClient
from msgraph.generated.models.user import User


@dataclass
class GraphContext:
    """Shared runtime state passed to every module.

    Attributes:
        graph_client: The authenticated Graph SDK client.
        is_app_only: True for client-credential (app-only) authentication,
            False for a delegated user token.
        region: Search region for the Graph Search API, e.g. `EMEA`.
        cached_user: The signed-in user from `/me`; delegated auth only.
        token_scopes: Scopes granted to a delegated token; empty for
            app-only.
        json_output: Whether `--json` was given; stdout must then carry only
            the final JSON.
        ndjson_output: Whether `--ndjson` was given; stdout carries one JSON
            object per line.
        token_getter: Async callable returning the current bearer token,
            for raw HTTP requests, or `None`.
        raw_identity: Identity string for cache partitioning when there is no
            cached user: the user OID (delegated) or `{tenant_id}:{client_id}`
            (app-only). Set by `cli.py`.
    """

    graph_client: GraphServiceClient
    is_app_only: bool
    region: str = "EMEA"
    cached_user: User | None = None
    token_scopes: frozenset[str] = field(default_factory=frozenset)
    json_output: bool = False
    ndjson_output: bool = False
    token_getter: Callable[[], Awaitable[str]] | None = field(
        default=None, repr=False, compare=False
    )
    raw_identity: str = field(default="", repr=False, compare=False)

    @property
    def identity_hash(self) -> str:
        """A 16-character hex SHA-256 prefix of the operator identity.

        Derived from the cached user's ID, else `raw_identity`, else the
        literal `unknown`. Used as the cache subdirectory so several tokens
        on one machine never overwrite each other's results.
        """
        raw = self.cached_user.id if self.cached_user and self.cached_user.id else self.raw_identity
        if not raw:
            return "unknown"
        return hashlib.sha256(raw.encode()).hexdigest()[:16]

    def has_scope(self, *required: str) -> bool:
        """Check whether the delegated token has every given scope.

        Args:
            *required: Scope names, e.g. `"Mail.Read"`.

        Returns:
            True if all of `required` are in `token_scopes`. Always False
            for app-only tokens unless `required` is empty.
        """
        return all(s in self.token_scopes for s in required)

    async def get_access_token(self) -> str | None:
        """Return the current bearer token from `token_getter`.

        Returns:
            The token, or `None` if no `token_getter` is set.
        """
        if self.token_getter is None:
            return None
        return await self.token_getter()

    @property
    def is_delegated(self) -> bool:
        """Whether authentication uses a delegated user token."""
        return not self.is_app_only
