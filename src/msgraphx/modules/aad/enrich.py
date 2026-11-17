# msgraphx/modules/aad/enrich.py
#
# Enriched user and group detail fetch via the SDK $batch endpoint.
# A single round-trip fans out to five or eight sub-requests and returns
# a merged dict that mirrors the shape GraphSpy expects.

from __future__ import annotations

import argparse
import json
from typing import Any, cast

# External library imports
from kiota_abstractions.base_request_configuration import RequestConfiguration
from loguru import logger
from msgraph.generated.users.item.user_item_request_builder import UserItemRequestBuilder
from msgraph_core.requests.batch_request_content import BatchRequestContent
from msgraph_core.requests.batch_response_content import BatchResponseContent

# Local library imports
from ...core.context import GraphContext
from ...utils import output
from ...utils.console import console
from ...utils.errors import ForbiddenGraphError, handle_graph_errors, raise_if_forbidden
from ...utils.roles import require_scopes

_USER_SELECT = [
    "displayName", "givenName", "surname", "userPrincipalName", "mail", "otherMails",
    "proxyAddresses", "mobilePhone", "businessPhones", "faxNumber",
    "createdDateTime", "lastPasswordChangeDateTime", "refreshTokensValidFromDateTime",
    "userType", "companyName", "jobTitle", "department", "officeLocation",
    "streetAddress", "city", "state", "country", "preferredLanguage",
    "id", "accountEnabled", "passwordPolicies", "licenseAssignmentStates",
    "creationType", "onPremisesSyncEnabled", "onPremisesDistinguishedName",
    "onPremisesSamAccountName", "onPremisesUserPrincipalName", "onPremisesDomainName",
    "onPremisesImmutableId", "onPremisesLastSyncDateTime", "onPremisesSecurityIdentifier",
    "securityIdentifier",
]

def _read_body(batch_resp: BatchResponseContent, req_id: str) -> dict[str, Any] | None:
    """Parse the JSON body of one batch sub-response.

    Args:
        batch_resp: The `$batch` response.
        req_id: The sub-request ID given when the batch was built.

    Returns:
        The decoded JSON body, or `None` if the sub-response is missing, has
        no body, or is not valid JSON.
    """
    item = batch_resp.get_response_by_id(req_id)
    if item is None or item.body is None:
        return None
    try:
        item.body.seek(0)
        return json.loads(item.body.read())
    except Exception:
        return None

async def _send_batch(
    context: GraphContext, batch_content: BatchRequestContent
) -> BatchResponseContent:
    """Send a `$batch` request.

    Args:
        context: Shared Graph runtime context.
        batch_content: The sub-requests to send in one round-trip.

    Returns:
        The batch response holding every sub-response.

    Raises:
        ForbiddenGraphError: Graph rejects the batch with 403.
    """
    try:
        # A single BatchRequestContent always yields a BatchResponseContent.
        return cast(
            BatchResponseContent,
            await context.graph_client.batch().post(batch_content),
        )
    except Exception as exc:
        raise_if_forbidden(exc)
        raise

async def fetch_user_details(context: GraphContext, user_id: str) -> dict[str, Any]:
    """Return a user's properties merged with their related collections.

    Fetches everything in one `$batch` round-trip. Related collections the
    token cannot read are returned empty, with a warning on 403.

    Args:
        context: Shared Graph runtime context.
        user_id: User object ID or UPN.

    Returns:
        The user's raw JSON properties (including `transitiveMemberOf`) plus
        `ownedObjects`, `ownedDevices`, `appRoleAssignments`, and
        `oauth2PermissionGrants` lists.

    Raises:
        ForbiddenGraphError: The token cannot read the user.
        RuntimeError: The user is not found or cannot be read for another
            reason.
    """
    user_builder = context.graph_client.users.by_user_id(user_id)
    UserQueryParams = UserItemRequestBuilder.UserItemRequestBuilderGetQueryParameters
    user_config = RequestConfiguration(
        query_parameters=UserQueryParams(
            expand=["transitiveMemberOf"],
            select=_USER_SELECT,
        )
    )

    batch_content = BatchRequestContent()
    batch_content.add_request_information(
        user_builder.to_get_request_information(user_config), request_id="userDetails"
    )
    batch_content.add_request_information(
        user_builder.owned_objects.to_get_request_information(), request_id="ownedObjects"
    )
    batch_content.add_request_information(
        user_builder.owned_devices.to_get_request_information(), request_id="ownedDevices"
    )
    batch_content.add_request_information(
        user_builder.app_role_assignments.to_get_request_information(),
        request_id="appRoleAssignments",
    )
    batch_content.add_request_information(
        user_builder.oauth2_permission_grants.to_get_request_information(),
        request_id="oauth2PermissionGrants",
    )

    batch_resp = await _send_batch(context, batch_content)
    status_codes = batch_resp.get_response_status_codes()

    base = _read_body(batch_resp, "userDetails")
    if not base:
        status = status_codes.get("userDetails", 0)
        if status == 403:

            raise ForbiddenGraphError(
                required="User.Read.All",
                granted=None,
                raw_message="Access denied reading user details.",
            )
        raise RuntimeError(f"User '{user_id}' not found or access denied (status {status}).")

    for key in ("ownedObjects", "ownedDevices", "appRoleAssignments", "oauth2PermissionGrants"):
        status = status_codes.get(key, 0)
        if status == 200:
            body = _read_body(batch_resp, key)
            base[key] = (body or {}).get("value", [])
        else:
            if status == 403:
                logger.warning(f"No access to {key} — skipped (insufficient permissions).")
            base[key] = []

    return base

async def fetch_group_details(context: GraphContext, group_id: str) -> dict[str, Any]:
    """Return a group's properties merged with its related resources.

    Fetches everything in one `$batch` round-trip. Related resources the
    token cannot read are returned empty, with a warning on 403 for
    collections.

    Args:
        context: Shared Graph runtime context.
        group_id: Group object ID.

    Returns:
        The group's raw JSON properties plus `transitiveMembers`, `owners`,
        `transitiveMemberOf`, `drives`, and `appRoleAssignments` lists and
        `team` and `sites` entries (`{}` when unavailable).

    Raises:
        ForbiddenGraphError: The token cannot read the group.
        RuntimeError: The group is not found or cannot be read for another
            reason.
    """
    group_builder = context.graph_client.groups.by_group_id(group_id)

    batch_content = BatchRequestContent()
    batch_content.add_request_information(
        group_builder.to_get_request_information(), request_id="groupDetails"
    )
    batch_content.add_request_information(
        group_builder.transitive_members.to_get_request_information(),
        request_id="transitiveMembers",
    )
    batch_content.add_request_information(
        group_builder.owners.to_get_request_information(), request_id="owners"
    )
    batch_content.add_request_information(
        group_builder.transitive_member_of.to_get_request_information(),
        request_id="transitiveMemberOf",
    )
    batch_content.add_request_information(
        group_builder.drives.to_get_request_information(), request_id="drives"
    )
    batch_content.add_request_information(
        group_builder.team.to_get_request_information(), request_id="team"
    )
    batch_content.add_request_information(
        group_builder.sites.to_get_request_information(), request_id="sites"
    )
    batch_content.add_request_information(
        group_builder.app_role_assignments.to_get_request_information(),
        request_id="appRoleAssignments",
    )

    batch_resp = await _send_batch(context, batch_content)
    status_codes = batch_resp.get_response_status_codes()

    base = _read_body(batch_resp, "groupDetails")
    if not base:
        status = status_codes.get("groupDetails", 0)
        if status == 403:

            raise ForbiddenGraphError(
                required="Group.Read.All",
                granted=None,
                raw_message="Access denied reading group details.",
            )
        raise RuntimeError(f"Group '{group_id}' not found or access denied (status {status}).")

    for key in ("transitiveMembers", "owners", "transitiveMemberOf", "drives",
                "appRoleAssignments"):
        status = status_codes.get(key, 0)
        if status == 200:
            body = _read_body(batch_resp, key)
            base[key] = (body or {}).get("value", [])
        else:
            if status == 403:
                logger.warning(f"No access to {key} — skipped (insufficient permissions).")
            base[key] = []

    # team and sites are singular objects, not collections
    for key in ("team", "sites"):
        status = status_codes.get(key, 0)
        if status == 200:
            body = _read_body(batch_resp, key)
            base[key] = (body or {}).get("value", body) if body else {}
        else:
            base[key] = {}

    return base

def add_arguments_user(parser: argparse.ArgumentParser) -> None:
    """Register the arguments for `aad user`.

    Args:
        parser: Parser for the `user` subcommand.
    """
    parser.add_argument("user_id", help="User object ID or UPN.")

def add_arguments_group(parser: argparse.ArgumentParser) -> None:
    """Register the arguments for `aad group`.

    Args:
        parser: Parser for the `group` subcommand.
    """
    parser.add_argument("group_id", help="Group object ID.")

@handle_graph_errors
@require_scopes("User.Read.All")
async def run_user(context: GraphContext, args: argparse.Namespace) -> int:
    """Show a user's enriched details.

    Prints JSON under `--json`, one NDJSON line under `--ndjson`, otherwise
    highlighted JSON. Warns if the token lacks `User.Read.All`.

    Args:
        context: Shared Graph runtime context.
        args: Parsed arguments; `args.user_id` is an object ID or UPN.

    Returns:
        0 on success.

    Raises:
        AuthenticationError: The Graph token is invalid or expired.
        ForbiddenGraphError: The token cannot read the user.
        RuntimeError: The user is not found.
    """
    logger.info(f"Fetching enriched user details: {args.user_id}")
    details = await fetch_user_details(context, args.user_id)

    if context.json_output:
        output.print_json(details)
        return 0

    if context.ndjson_output:
        output.print_ndjson_item(details)
        return 0

    console.print(f"[bold]User: {details.get('displayName', args.user_id)}[/bold]")
    console.rule()
    console.print_json(json.dumps(details, indent=2, default=str))
    return 0

@handle_graph_errors
@require_scopes("Group.Read.All")
async def run_group(context: GraphContext, args: argparse.Namespace) -> int:
    """Show a group's enriched details.

    Prints JSON under `--json`, one NDJSON line under `--ndjson`, otherwise
    highlighted JSON. Warns if the token lacks `Group.Read.All`.

    Args:
        context: Shared Graph runtime context.
        args: Parsed arguments; `args.group_id` is the group object ID.

    Returns:
        0 on success.

    Raises:
        AuthenticationError: The Graph token is invalid or expired.
        ForbiddenGraphError: The token cannot read the group.
        RuntimeError: The group is not found.
    """
    logger.info(f"Fetching enriched group details: {args.group_id}")
    details = await fetch_group_details(context, args.group_id)

    if context.json_output:
        output.print_json(details)
        return 0

    if context.ndjson_output:
        output.print_ndjson_item(details)
        return 0

    console.print(f"[bold]Group: {details.get('displayName', args.group_id)}[/bold]")
    console.rule()
    console.print_json(json.dumps(details, indent=2, default=str))
    return 0
