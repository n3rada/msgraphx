# msgraphx/modules/aad/enrich.py
#
# Enriched user and group detail fetch via the $batch endpoint.
# A single round-trip fans out to five or eight sub-requests and returns
# a merged dict that mirrors the shape GraphSpy expects.
#
# Required delegated permissions:
#   User.Read.All, Group.Read.All, GroupMember.Read.All

from __future__ import annotations

import argparse
import json
from urllib.parse import quote_plus

# External library imports
import httpx
from loguru import logger
from rich.table import Table

# Local library imports
from ...core.context import GraphContext
from ...utils import output
from ...utils.console import console
from ...utils.errors import handle_graph_errors

_BATCH_URL = "https://graph.microsoft.com/v1.0/$batch"

_USER_SELECT = (
    "displayName,givenName,surname,userPrincipalName,mail,otherMails,"
    "proxyAddresses,mobilePhone,businessPhones,faxNumber,"
    "createdDateTime,lastPasswordChangeDateTime,refreshTokensValidFromDateTime,"
    "userType,companyName,jobTitle,department,officeLocation,"
    "streetAddress,city,state,country,preferredLanguage,"
    "id,accountEnabled,passwordPolicies,licenseAssignmentStates,"
    "creationType,onPremisesSyncEnabled,onPremisesDistinguishedName,"
    "onPremisesSamAccountName,onPremisesUserPrincipalName,onPremisesDomainName,"
    "onPremisesImmutableId,onPremisesLastSyncDateTime,onPremisesSecurityIdentifier,"
    "securityIdentifier"
)


async def _batch(token: str, requests: list[dict]) -> list[dict]:
    async with httpx.AsyncClient(verify=False, timeout=30) as client:
        resp = await client.post(
            _BATCH_URL,
            json={"requests": requests},
            headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
        )
    if resp.status_code != 200:
        raise RuntimeError(f"$batch failed: {resp.status_code} {resp.text[:200]}")
    throttled = [r for r in resp.json()["responses"] if r["status"] == 429]
    if throttled:
        raise RuntimeError("$batch request was throttled (429). Retry later.")
    return resp.json()["responses"]


async def fetch_user_details(context: GraphContext, user_id: str) -> dict:
    """Return a merged dict with full user properties and related collections."""
    token = await context.get_access_token()
    if not token:
        raise RuntimeError("No access token available.")

    uid = quote_plus(user_id)
    responses = await _batch(token, [
        {
            "id": "userDetails",
            "method": "GET",
            "url": f"/users/{uid}?$expand=transitiveMemberOf&$select={_USER_SELECT}",
        },
        {"id": "ownedObjects", "method": "GET", "url": f"/users/{uid}/ownedObjects"},
        {"id": "ownedDevices", "method": "GET", "url": f"/users/{uid}/ownedDevices"},
        {"id": "appRoleAssignments", "method": "GET", "url": f"/users/{uid}/appRoleAssignments"},
        {"id": "oauth2PermissionGrants", "method": "GET", "url": f"/users/{uid}/oauth2PermissionGrants"},
    ])

    base = next((r["body"] for r in responses if r["id"] == "userDetails" and r["status"] == 200), None)
    if base is None:
        raise RuntimeError(f"User '{user_id}' not found or access denied.")

    for r in responses:
        if r["id"] == "userDetails":
            continue
        base[r["id"]] = r["body"].get("value", []) if r["status"] == 200 else []

    return base


async def fetch_group_details(context: GraphContext, group_id: str) -> dict:
    """Return a merged dict with full group properties and related collections."""
    token = await context.get_access_token()
    if not token:
        raise RuntimeError("No access token available.")

    gid = quote_plus(group_id)
    responses = await _batch(token, [
        {"id": "groupDetails", "method": "GET", "url": f"/groups/{gid}"},
        {"id": "transitiveMembers", "method": "GET", "url": f"/groups/{gid}/transitiveMembers"},
        {"id": "owners", "method": "GET", "url": f"/groups/{gid}/owners"},
        {"id": "transitiveMemberOf", "method": "GET", "url": f"/groups/{gid}/transitiveMemberOf"},
        {"id": "drives", "method": "GET", "url": f"/groups/{gid}/drives"},
        {"id": "team", "method": "GET", "url": f"/groups/{gid}/team"},
        {"id": "sites", "method": "GET", "url": f"/groups/{gid}/sites"},
        {"id": "appRoleAssignments", "method": "GET", "url": f"/groups/{gid}/appRoleAssignments"},
    ])

    base = next((r["body"] for r in responses if r["id"] == "groupDetails" and r["status"] == 200), None)
    if base is None:
        raise RuntimeError(f"Group '{group_id}' not found or access denied.")

    for r in responses:
        if r["id"] == "groupDetails":
            continue
        if r["status"] == 200:
            body = r["body"]
            base[r["id"]] = body.get("value", body)
        else:
            base[r["id"]] = []

    return base


def add_arguments_user(parser: argparse.ArgumentParser) -> None:
    parser.add_argument("user_id", help="User object ID or UPN.")


def add_arguments_group(parser: argparse.ArgumentParser) -> None:
    parser.add_argument("group_id", help="Group object ID.")


@handle_graph_errors
async def run_user(context: GraphContext, args: argparse.Namespace) -> int:
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
async def run_group(context: GraphContext, args: argparse.Namespace) -> int:
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
