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
from typing import Any

# Local library imports
from .core.context import GraphContext
from .core.auth import create_context
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
from .modules.outlook import contacts as _outlook_contacts
from .modules.outlook import search as _outlook_search
from .modules.sharepoint import search as _sp_search
from .modules.sharepoint import sites as _sp_sites
from .modules.teams import channel as _teams_channel
from .modules.teams import chat as _teams_chat
from .modules.teams import meetings as _teams_meetings
from .utils.tokens import TokenManager


class _MeNamespace:
    """Signed-in user resources, exposed as `Session.me` (delegated only)."""

    def __init__(self, context: GraphContext) -> None:
        """Bind the namespace to a context.

        Args:
            context: The session's Graph context.
        """
        self._ctx = context

    async def people(
        self, top: int = 25, search: str | None = None
    ) -> list[dict[str, Any]]:
        """Return the people most relevant to the user.

        Args:
            top: Maximum number of people, capped at 1000.
            search: Name to search for, or `None` for the full ranking.

        Returns:
            One record per person with name, job, contact, and office fields.
        """
        return await _me_people.fetch(self._ctx, top=top, search=search)

    async def calendar(
        self,
        top: int = 50,
        after: str | None = None,
        before: str | None = None,
    ) -> list[dict[str, Any]]:
        """Return the user's calendar events.

        Args:
            top: Maximum number of events.
            after: ISO 8601 datetime; keep events starting at or after it.
            before: ISO 8601 datetime; keep events ending at or before it.

        Returns:
            One record per event with subject, times, organizer, location,
            attendees, and links.
        """
        return await _me_calendar.fetch(self._ctx, top=top, after=after, before=before)

    async def groups(self, visibility: str | None = None) -> list[dict[str, Any]]:
        """Return the groups the user belongs to, including nested ones.

        Args:
            visibility: `Private` or `Public` to filter, or `None` for all.

        Returns:
            One record per group.
        """
        return await _me_groups.fetch(self._ctx, visibility=visibility)

    async def planner(self, top: int = 50) -> list[dict[str, Any]]:
        """Return the Planner tasks assigned to the user.

        Args:
            top: Maximum number of tasks.

        Returns:
            One record per task.
        """
        return await _me_planner.fetch(self._ctx, top=top)

    async def shared(self, top: int = 25) -> list[dict[str, Any]]:
        """Return documents shared with the user.

        Args:
            top: Maximum number of documents.

        Returns:
            One record per document.
        """
        return await _me_shared.fetch(self._ctx, top=top)

    async def used(self, top: int = 25) -> list[dict[str, Any]]:
        """Return documents the user recently used.

        Args:
            top: Maximum number of documents.

        Returns:
            One record per document.
        """
        return await _me_used.fetch(self._ctx, top=top)

    async def trending(
        self, top: int = 25, type_filter: str | None = None
    ) -> list[dict[str, Any]]:
        """Return documents trending around the user.

        Args:
            top: Maximum number of documents, capped at 100.
            type_filter: Case-insensitive substring of the resource type,
                e.g. `Excel`, or `None` for all types.

        Returns:
            One record per document, ranked.
        """
        return await _me_trending.fetch(self._ctx, top=top, type_filter=type_filter)


class _AadNamespace:
    """Entra ID directory queries, exposed as `Session.aad`."""

    def __init__(self, context: GraphContext) -> None:
        """Bind the namespace to a context.

        Args:
            context: The session's Graph context.
        """
        self._ctx = context

    async def users(
        self, query: str = "*", odata_filter: str | None = None
    ) -> list[dict[str, Any]]:
        """Return users matching a query.

        Args:
            query: Text searched across display name, UPN, and mail.
            odata_filter: OData `$filter` expression that replaces `query`.

        Returns:
            One record per user.
        """
        return await fetch_users(self._ctx, query=query, odata_filter=odata_filter)

    async def groups(
        self,
        query: str = "*",
        contains: bool = False,
        synced_only: bool = False,
        odata_filter: str | None = None,
    ) -> list[dict[str, Any]]:
        """Return groups matching a query.

        Args:
            query: Text to search for in group names.
            contains: Match `query` anywhere in the display name or mail
                nickname, client side, after fetching every group.
            synced_only: Keep only groups synced from on-premises AD.
            odata_filter: OData `$filter` expression that replaces `query`
                and `contains`.

        Returns:
            One record per group.
        """
        return await fetch_groups(
            self._ctx,
            query=query,
            contains=contains,
            synced_only=synced_only,
            odata_filter=odata_filter,
        )

    async def devices(self, query: str = "*") -> list[dict[str, Any]]:
        """Return devices whose display name matches a query.

        Args:
            query: Text to search for.

        Returns:
            One record per device.
        """
        return await fetch_devices(self._ctx, query=query)

    async def service_principals(self, query: str = "*") -> list[dict[str, Any]]:
        """Return service principals whose display name matches a query.

        Args:
            query: Text to search for.

        Returns:
            One record per service principal.
        """
        return await fetch_service_principals(self._ctx, query=query)

    async def applications(self, query: str = "*") -> list[dict[str, Any]]:
        """Return app registrations whose display name matches a query.

        Args:
            query: Text to search for.

        Returns:
            One record per application.
        """
        return await fetch_applications(self._ctx, query=query)

    async def ca_policies(self, state: str | None = None) -> list[dict[str, Any]]:
        """Return the tenant's conditional access policies.

        Args:
            state: `enabled`, `disabled`, or `report` (report-only) to
                filter, or `None` for every policy.

        Returns:
            One record per policy.
        """
        return await _aad_ca.fetch(self._ctx, state=state)

    async def role_assignments(
        self, odata_filter: str | None = None
    ) -> list[dict[str, Any]]:
        """Return directory role assignments with role and principal names.

        Args:
            odata_filter: OData `$filter` expression, e.g.
                `"principalId eq 'GUID'"`, or `None` for every assignment.

        Returns:
            One record per assignment.
        """
        return await _aad_roles.fetch(self._ctx, odata_filter=odata_filter)

    async def user_details(self, user_id: str) -> dict[str, Any]:
        """Return a user's properties merged with their related collections.

        Args:
            user_id: User object ID or UPN.

        Returns:
            The user's properties plus owned objects, devices, and grants.
        """
        return await _aad_enrich.fetch_user_details(self._ctx, user_id)

    async def group_details(self, group_id: str) -> dict[str, Any]:
        """Return a group's properties merged with its related resources.

        Args:
            group_id: Group object ID.

        Returns:
            The group's properties plus members, owners, and resources.
        """
        return await _aad_enrich.fetch_group_details(self._ctx, group_id)

    async def pim_roles(self) -> list[dict[str, Any]]:
        """Return PIM role assignments.

        Returns:
            One record per assignment.
        """
        return await _aad_pim.fetch(self._ctx)



class _OutlookNamespace:
    """Mailbox queries, exposed as `Session.outlook` (delegated only)."""

    def __init__(self, context: GraphContext) -> None:
        """Bind the namespace to a context.

        Args:
            context: The session's Graph context.
        """
        self._ctx = context

    async def search(
        self,
        query: str = "*",
        from_addr: str | None = None,
        subject: str | None = None,
        has_attachments: bool = False,
        after: str | None = None,
        before: str | None = None,
    ) -> list[dict[str, Any]]:
        """Search the user's mail.

        Args:
            query: KQL query string; `*` matches everything.
            from_addr: Keep only mail from this sender.
            subject: Keep only mail whose subject matches this text.
            has_attachments: Keep only mail with attachments.
            after: ISO 8601 timestamp; keep mail received at or after it.
            before: ISO 8601 timestamp; keep mail received at or before it.

        Returns:
            One record per message.
        """
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
    ) -> dict[str, list[dict[str, Any]]]:
        """Rank the people the user exchanges mail with most.

        Args:
            after: ISO 8601 timestamp; include mail at or after it.
            before: ISO 8601 timestamp; include mail at or before it.
            only: `sent` or `received` to analyse one direction, or `None`
                for both.
            top: Maximum contacts per direction, or `None` for all.

        Returns:
            A dict with `sent` and `received` lists of ranked contacts.
        """
        return await _outlook_contacts.fetch(self._ctx, after=after, before=before, only=only, top=top)


class _SharePointNamespace:
    """SharePoint and OneDrive queries, exposed as `Session.sharepoint`."""

    def __init__(self, context: GraphContext) -> None:
        """Bind the namespace to a context.

        Args:
            context: The session's Graph context.
        """
        self._ctx = context

    async def search(
        self,
        query: str = "*",
        after: str | None = None,
        before: str | None = None,
        group_ids: list[str] | None = None,
        region: str | None = None,
        drive_id: str | None = None,
    ) -> list[dict[str, Any]]:
        """Search SharePoint and OneDrive files.

        Args:
            query: KQL query string; `*` matches everything.
            after: ISO 8601 date; keep items after it.
            before: ISO 8601 date; keep items before it.
            group_ids: Restrict results to sites of these M365 group IDs.
            region: Accepted for compatibility; unused.
            drive_id: Restrict results to this drive.

        Returns:
            One record per matching drive item.
        """
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
        """Return the SharePoint sites accessible to the caller.

        Args:
            show_public: Include public sites.

        Returns:
            The accessible sites.
        """
        return await _sp_sites.fetch(self._ctx, show_public=show_public)


class _TeamsNamespace:
    """Microsoft Teams queries, exposed as `Session.teams` (delegated only)."""

    def __init__(self, context: GraphContext) -> None:
        """Bind the namespace to a context.

        Args:
            context: The session's Graph context.
        """
        self._ctx = context

    async def chats(self, top: int = 20) -> list[dict[str, Any]]:
        """Return the user's most recently active chats.

        Args:
            top: Maximum number of chats.

        Returns:
            One record per chat.
        """
        return await _teams_chat.fetch_chats(self._ctx, top=top)

    async def chat_search(
        self,
        query: str = "*",
        from_addr: str | None = None,
        after: str | None = None,
        before: str | None = None,
    ) -> list[dict[str, Any]]:
        """Search Teams chat messages.

        Args:
            query: KQL query string; `*` matches everything.
            from_addr: Keep only messages from this sender.
            after: ISO 8601 timestamp; keep messages at or after it.
            before: ISO 8601 timestamp; keep messages at or before it.

        Returns:
            One record per message.
        """
        return await _teams_chat.fetch_search(self._ctx, query=query, from_addr=from_addr, after=after, before=before)

    async def joined_teams(self) -> list[dict[str, Any]]:
        """Return the teams the user has joined.

        Returns:
            One record per team.
        """
        return await _teams_channel.fetch_teams(self._ctx)

    async def channel_search(
        self,
        query: str = "*",
        from_addr: str | None = None,
        after: str | None = None,
        before: str | None = None,
    ) -> list[dict[str, Any]]:
        """Search Teams channel messages.

        Args:
            query: KQL query string; `*` matches everything.
            from_addr: Keep only messages from this sender.
            after: ISO 8601 timestamp; keep messages at or after it.
            before: ISO 8601 timestamp; keep messages at or before it.

        Returns:
            One record per message.
        """
        return await _teams_channel.fetch_search(self._ctx, query=query, from_addr=from_addr, after=after, before=before)

    async def meetings(self, top: int = 25) -> list[dict[str, Any]]:
        """Return the user's online meetings.

        Args:
            top: Maximum number of meetings.

        Returns:
            One record per meeting.
        """
        return await _teams_meetings.fetch(self._ctx, top=top)

    async def transcripts(self, meeting_id: str) -> list[dict[str, Any]]:
        """Return every transcript of an online meeting.

        Args:
            meeting_id: The online meeting ID.

        Returns:
            One record per transcript.
        """
        return await _teams_meetings.fetch_transcripts(self._ctx, meeting_id=meeting_id)


class Session:
    """An authenticated msgraphx session for library use.

    Methods return plain dicts and raise on API errors; callers handle the
    exceptions. The underlying `GraphContext` is private.

    Example:
        session = await Session.create(access_token="eyJ...")
        people = await session.me.people(top=25)
        admins = await session.aad.users(query="admin")
        files = await session.sharepoint.search("filetype:xlsx credentials")

    Attributes:
        me: Signed-in user resources (delegated only).
        aad: Entra ID directory queries.
        outlook: Mailbox queries (delegated only).
        sharepoint: SharePoint and OneDrive queries.
        teams: Microsoft Teams queries (delegated only).
    """

    def __init__(
        self,
        context: GraphContext,
        token_manager: TokenManager | None = None,
    ) -> None:
        """Wrap an existing context; prefer `Session.create`.

        Args:
            context: An authenticated Graph context.
            token_manager: Manager for the delegated token, used by
                `start_refresh`, or `None`.
        """
        self._ctx = context
        self._token_manager = token_manager
        self.me = _MeNamespace(context)
        self.aad = _AadNamespace(context)
        self.outlook = _OutlookNamespace(context)
        self.sharepoint = _SharePointNamespace(context)
        self.teams = _TeamsNamespace(context)

    @property
    def is_app_only(self) -> bool:
        """Whether the session uses app-only (client credential) auth."""
        return self._ctx.is_app_only

    @property
    def region(self) -> str:
        """The Graph Search API region, e.g. `EMEA`."""
        return self._ctx.region

    def start_refresh(self) -> threading.Thread | None:
        """Start a daemon thread that keeps the delegated token fresh.

        Call once after `Session.create` for long-running work. Refreshed
        tokens are persisted to `.roadtools_auth`.

        Returns:
            The refresh thread, or `None` if there is no token manager or no
            refresh token.
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
    ) -> Session:
        """Build an authenticated session.

        Example:
            session = await Session.create(
                access_token="eyJ...", refresh_token="0.A..."
            )
            session = await Session.create(
                tenant_id="...", client_id="...", client_secret="..."
            )

        Args:
            access_token: Delegated JWT access token for Graph.
            refresh_token: Refresh token enabling `start_refresh`.
            tenant_id: Tenant ID for app-only authentication.
            client_id: Application (client) ID for app-only authentication.
            client_secret: Client secret for app-only authentication.
            region: Search region for the Graph Search API.

        Returns:
            The session. For delegated auth it holds a token manager when the
            access token can be decoded.

        Raises:
            ValueError: If the credentials are incomplete, or the access
                token is malformed or expired.
        """
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
