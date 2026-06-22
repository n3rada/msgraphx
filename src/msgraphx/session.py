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

# Local library imports
from .core.context import GraphContext
from .modules.aad import ca as _aad_ca
from .modules.aad import enrich as _aad_enrich
from .modules.aad import pim as _aad_pim
from .modules.aad import roles as _aad_roles
from .modules.aad.search import (
    fetch_applications,
    fetch_devices,
    fetch_groups,
    fetch_service_principals,
    fetch_users,
)
from .modules.me import calendar as _me_calendar
from .modules.me import groups as _me_groups
from .modules.me import people as _me_people
from .modules.me import planner as _me_planner
from .modules.me import shared as _me_shared
from .modules.me import trending as _me_trending
from .modules.me import used as _me_used
from .modules.mfa import security_info as _mfa_security_info
from .modules.outlook import contacts as _outlook_contacts
from .modules.outlook import search as _outlook_search
from .modules.sharepoint import search as _sp_search
from .modules.sharepoint import sites as _sp_sites
from .modules.teams import channel as _teams_channel
from .modules.teams import chat as _teams_chat
from .modules.teams import meetings as _teams_meetings
from .utils.tokens import TokenManager


class _MeNamespace:
    def __init__(self, context: GraphContext) -> None:
        self._ctx = context

    async def people(self, top: int = 25, search: str | None = None) -> list[dict]:
        return await _me_people.fetch(self._ctx, top=top, search=search)

    async def calendar(
        self,
        top: int = 50,
        after: str | None = None,
        before: str | None = None,
    ) -> list[dict]:
        return await _me_calendar.fetch(self._ctx, top=top, after=after, before=before)

    async def groups(self, visibility: str | None = None) -> list[dict]:
        return await _me_groups.fetch(self._ctx, visibility=visibility)

    async def planner(self, top: int = 50) -> list[dict]:
        return await _me_planner.fetch(self._ctx, top=top)

    async def shared(self, top: int = 25) -> list[dict]:
        return await _me_shared.fetch(self._ctx, top=top)

    async def used(self, top: int = 25) -> list[dict]:
        return await _me_used.fetch(self._ctx, top=top)

    async def trending(self, top: int = 25, type_filter: str | None = None) -> list[dict]:
        return await _me_trending.fetch(self._ctx, top=top, type_filter=type_filter)


class _AadNamespace:
    def __init__(self, context: GraphContext) -> None:
        self._ctx = context

    async def users(self, query: str = "*", odata_filter: str | None = None) -> list[dict]:
        return await fetch_users(self._ctx, query=query, odata_filter=odata_filter)

    async def groups(
        self,
        query: str = "*",
        contains: bool = False,
        synced_only: bool = False,
        odata_filter: str | None = None,
    ) -> list[dict]:
        return await fetch_groups(
            self._ctx,
            query=query,
            contains=contains,
            synced_only=synced_only,
            odata_filter=odata_filter,
        )

    async def devices(self, query: str = "*") -> list[dict]:
        return await fetch_devices(self._ctx, query=query)

    async def service_principals(self, query: str = "*") -> list[dict]:
        return await fetch_service_principals(self._ctx, query=query)

    async def applications(self, query: str = "*") -> list[dict]:
        return await fetch_applications(self._ctx, query=query)

    async def ca_policies(self, state: str | None = None) -> list[dict]:
        return await _aad_ca.fetch(self._ctx, state=state)

    async def role_assignments(self, odata_filter: str | None = None) -> list[dict]:
        return await _aad_roles.fetch(self._ctx, odata_filter=odata_filter)

    async def user_details(self, user_id: str) -> dict:
        return await _aad_enrich.fetch_user_details(self._ctx, user_id)

    async def group_details(self, group_id: str) -> dict:
        return await _aad_enrich.fetch_group_details(self._ctx, group_id)

    async def pim_roles(self) -> list[dict]:
        return await _aad_pim.fetch(self._ctx)


class _MfaNamespace:
    """MFA manipulation via mysignins.microsoft.com.

    Requires a dedicated access token scoped to resource
    19db86c3-b2b9-44cc-b339-36da233a3be2 (the My Sign-Ins portal),
    which is different from the Graph API token.
    """

    async def available_methods(self, access_token: str) -> list[dict]:
        return await _mfa_security_info.available_methods(access_token)

    async def add_otp_backdoor(self, access_token: str) -> str | None:
        return await _mfa_security_info.add_otp_backdoor(access_token)

    async def add_phone(
        self,
        access_token: str,
        country_code: str,
        phone_number: str,
        phone_type: str = "sms",
    ) -> dict:
        return await _mfa_security_info.add_phone(access_token, country_code, phone_number, phone_type)

    async def add_email(self, access_token: str, email: str) -> dict:
        return await _mfa_security_info.add_email(access_token, email)

    async def verify(
        self,
        access_token: str,
        security_info_type: int,
        verification_context: str | None,
        verification_data: str,
    ) -> dict:
        return await _mfa_security_info.verify(access_token, security_info_type, verification_context, verification_data)

    async def delete(self, access_token: str, security_info_type: int, data: str) -> dict:
        return await _mfa_security_info.delete(access_token, security_info_type, data)


class _OutlookNamespace:
    def __init__(self, context: GraphContext) -> None:
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
        return await _outlook_search.fetch(
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
        return await _outlook_contacts.fetch(self._ctx, after=after, before=before, only=only, top=top)


class _SharePointNamespace:
    def __init__(self, context: GraphContext) -> None:
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
        return await _sp_search.fetch(
            self._ctx,
            query=query,
            after=after,
            before=before,
            group_ids=group_ids,
            region=region,
            drive_id=drive_id,
        )

    async def sites(self, show_public: bool = False) -> dict:
        return await _sp_sites.fetch(self._ctx, show_public=show_public)


class _TeamsNamespace:
    def __init__(self, context: GraphContext) -> None:
        self._ctx = context

    async def chats(self, top: int = 20) -> list[dict]:
        return await _teams_chat.fetch_chats(self._ctx, top=top)

    async def chat_search(
        self,
        query: str = "*",
        from_addr: str | None = None,
        after: str | None = None,
        before: str | None = None,
    ) -> list[dict]:
        return await _teams_chat.fetch_search(self._ctx, query=query, from_addr=from_addr, after=after, before=before)

    async def joined_teams(self) -> list[dict]:
        return await _teams_channel.fetch_teams(self._ctx)

    async def channel_search(
        self,
        query: str = "*",
        from_addr: str | None = None,
        after: str | None = None,
        before: str | None = None,
    ) -> list[dict]:
        return await _teams_channel.fetch_search(self._ctx, query=query, from_addr=from_addr, after=after, before=before)

    async def meetings(self, top: int = 25) -> list[dict]:
        return await _teams_meetings.fetch(self._ctx, top=top)

    async def transcripts(self, meeting_id: str) -> list[dict]:
        return await _teams_meetings.fetch_transcripts(self._ctx, meeting_id=meeting_id)


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
        context: GraphContext,
        token_manager: TokenManager | None = None,
    ) -> None:
        self._ctx = context
        self._token_manager = token_manager
        self.me = _MeNamespace(context)
        self.aad = _AadNamespace(context)
        self.outlook = _OutlookNamespace(context)
        self.sharepoint = _SharePointNamespace(context)
        self.teams = _TeamsNamespace(context)
        self.mfa = _MfaNamespace()

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
        # Lazy: __init__.py imports Session at module level, so importing
        # create_context here avoids a circular import at load time.
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
            try:
                token_manager = TokenManager(access_token, refresh_token)
            except Exception:
                pass

        return cls(ctx, token_manager=token_manager)
