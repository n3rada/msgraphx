# msgraphx/session.py
#
# High-level library API. Session wraps a GraphContext and exposes
# service namespaces so callers never need to import internal modules:
#
#   session = await Session.create(access_token="eyJ...")
#   people  = await session.me.people(top=25)
#   users   = await session.aad.users(query="admin")
#   files   = await session.sharepoint.search("filetype:xlsx")

from __future__ import annotations

import threading
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from .core.context import GraphContext
    from .utils.tokens import TokenManager


class _MeNamespace:
    def __init__(self, context: "GraphContext") -> None:
        self._ctx = context

    async def people(self, top: int = 25, search: str | None = None) -> list[dict]:
        from .modules.me.people import fetch
        return await fetch(self._ctx, top=top, search=search)

    async def calendar(
        self,
        top: int = 50,
        after: str | None = None,
        before: str | None = None,
    ) -> list[dict]:
        from .modules.me.calendar import fetch
        return await fetch(self._ctx, top=top, after=after, before=before)

    async def groups(self, visibility: str | None = None) -> list[dict]:
        from .modules.me.groups import fetch
        return await fetch(self._ctx, visibility=visibility)

    async def planner(self, top: int = 50) -> list[dict]:
        from .modules.me.planner import fetch
        return await fetch(self._ctx, top=top)

    async def shared(self, top: int = 25) -> list[dict]:
        from .modules.me.shared import fetch
        return await fetch(self._ctx, top=top)

    async def used(self, top: int = 25) -> list[dict]:
        from .modules.me.used import fetch
        return await fetch(self._ctx, top=top)

    async def trending(self, top: int = 25, type_filter: str | None = None) -> list[dict]:
        from .modules.me.trending import fetch
        return await fetch(self._ctx, top=top, type_filter=type_filter)


class _AadNamespace:
    def __init__(self, context: "GraphContext") -> None:
        self._ctx = context

    async def users(
        self,
        query: str = "*",
        odata_filter: str | None = None,
    ) -> list[dict]:
        from .modules.aad.search import fetch_users
        return await fetch_users(self._ctx, query=query, odata_filter=odata_filter)

    async def groups(
        self,
        query: str = "*",
        contains: bool = False,
        synced_only: bool = False,
        odata_filter: str | None = None,
    ) -> list[dict]:
        from .modules.aad.search import fetch_groups
        return await fetch_groups(
            self._ctx,
            query=query,
            contains=contains,
            synced_only=synced_only,
            odata_filter=odata_filter,
        )

    async def devices(self, query: str = "*") -> list[dict]:
        from .modules.aad.search import fetch_devices
        return await fetch_devices(self._ctx, query=query)

    async def service_principals(self, query: str = "*") -> list[dict]:
        from .modules.aad.search import fetch_service_principals
        return await fetch_service_principals(self._ctx, query=query)

    async def applications(self, query: str = "*") -> list[dict]:
        from .modules.aad.search import fetch_applications
        return await fetch_applications(self._ctx, query=query)

    async def ca_policies(self, state: str | None = None) -> list[dict]:
        from .modules.aad.ca import fetch
        return await fetch(self._ctx, state=state)

    async def role_assignments(self, odata_filter: str | None = None) -> list[dict]:
        from .modules.aad.roles import fetch
        return await fetch(self._ctx, odata_filter=odata_filter)


class _OutlookNamespace:
    def __init__(self, context: "GraphContext") -> None:
        self._ctx = context

    async def search(
        self,
        query: str = "*",
        from_addr: str | None = None,
        subject: str | None = None,
        has_attachments: bool = False,
        after: str | None = None,
        before: str | None = None,
    ) -> list[dict]:
        from .modules.outlook.search import fetch
        return await fetch(
            self._ctx,
            query=query,
            from_addr=from_addr,
            subject=subject,
            has_attachments=has_attachments,
            after=after,
            before=before,
        )

    async def contacts(
        self,
        after: str | None = None,
        before: str | None = None,
        only: str | None = None,
        top: int | None = None,
    ) -> dict:
        from .modules.outlook.contacts import fetch
        return await fetch(self._ctx, after=after, before=before, only=only, top=top)


class _SharePointNamespace:
    def __init__(self, context: "GraphContext") -> None:
        self._ctx = context

    async def search(
        self,
        query: str = "*",
        after: str | None = None,
        before: str | None = None,
        group_ids: list[str] | None = None,
        region: str | None = None,
        drive_id: str | None = None,
    ) -> list[dict]:
        from .modules.sharepoint.search import fetch
        return await fetch(
            self._ctx,
            query=query,
            after=after,
            before=before,
            group_ids=group_ids,
            region=region,
            drive_id=drive_id,
        )

    async def sites(self, show_public: bool = False) -> dict:
        from .modules.sharepoint.sites import fetch
        return await fetch(self._ctx, show_public=show_public)


class _TeamsNamespace:
    def __init__(self, context: "GraphContext") -> None:
        self._ctx = context

    async def chats(self, top: int = 20) -> list[dict]:
        from .modules.teams.chat import fetch_chats
        return await fetch_chats(self._ctx, top=top)

    async def chat_search(
        self,
        query: str = "*",
        from_addr: str | None = None,
        after: str | None = None,
        before: str | None = None,
    ) -> list[dict]:
        from .modules.teams.chat import fetch_search
        return await fetch_search(self._ctx, query=query, from_addr=from_addr, after=after, before=before)

    async def joined_teams(self) -> list[dict]:
        from .modules.teams.channel import fetch_teams
        return await fetch_teams(self._ctx)

    async def channel_search(
        self,
        query: str = "*",
        from_addr: str | None = None,
        after: str | None = None,
        before: str | None = None,
    ) -> list[dict]:
        from .modules.teams.channel import fetch_search
        return await fetch_search(self._ctx, query=query, from_addr=from_addr, after=after, before=before)

    async def meetings(self, top: int = 25) -> list[dict]:
        from .modules.teams.meetings import fetch
        return await fetch(self._ctx, top=top)

    async def transcripts(self, meeting_id: str) -> list[dict]:
        from .modules.teams.meetings import fetch_transcripts
        return await fetch_transcripts(self._ctx, meeting_id=meeting_id)


class Session:
    """Library entry point for msgraphx.

    All methods return plain dicts and raise on API errors; callers handle
    exceptions. The underlying GraphContext is private; external code should
    never need to import it.

    Usage::

        session = await Session.create(access_token="eyJ...")
        people  = await session.me.people(top=25)
        admins  = await session.aad.users(query="admin")
        files   = await session.sharepoint.search("filetype:xlsx credentials")
        events  = await session.me.calendar(after="2025-01-01T00:00:00")

    App-only (client credentials)::

        session = await Session.create(
            tenant_id="...", client_id="...", client_secret="..."
        )
    """

    def __init__(
        self,
        context: "GraphContext",
        token_manager: "TokenManager | None" = None,
    ) -> None:
        self._ctx = context
        self._token_manager = token_manager
        self.me = _MeNamespace(context)
        self.aad = _AadNamespace(context)
        self.outlook = _OutlookNamespace(context)
        self.sharepoint = _SharePointNamespace(context)
        self.teams = _TeamsNamespace(context)

    @property
    def is_app_only(self) -> bool:
        return self._ctx.is_app_only

    @property
    def region(self) -> str:
        return self._ctx.region

    def start_refresh(self) -> threading.Thread | None:
        """Start a background token refresh thread.

        Returns the thread if a refresh token is available, None otherwise.
        Call once after Session.create() to keep delegated sessions alive::

            session = await Session.create(access_token="...", refresh_token="...")
            session.start_refresh()
        """
        if self._token_manager is None:
            return None
        return self._token_manager.start_auto_refresh()

    @classmethod
    async def create(
        cls,
        access_token: str | None = None,
        refresh_token: str | None = None,
        tenant_id: str | None = None,
        client_id: str | None = None,
        client_secret: str | None = None,
        region: str = "EMEA",
    ) -> "Session":
        """Build an authenticated Session.

        Delegated::

            session = await Session.create(access_token="eyJ...", refresh_token="0.A...")
            session.start_refresh()  # keep alive for multi-hour ops

        App-only::

            session = await Session.create(
                tenant_id="...", client_id="...", client_secret="..."
            )
        """
        from . import create_context

        ctx = await create_context(
            access_token=access_token,
            refresh_token=refresh_token,
            tenant_id=tenant_id,
            client_id=client_id,
            client_secret=client_secret,
            region=region,
        )

        token_manager = None
        if access_token and not tenant_id:
            from .utils.tokens import TokenManager
            try:
                token_manager = TokenManager(access_token, refresh_token)
            except Exception:
                pass

        return cls(ctx, token_manager=token_manager)
