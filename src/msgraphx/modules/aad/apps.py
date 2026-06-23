# msgraphx/modules/aad/apps.py
#
# aad app  <id>  — enriched single-app profile (registration + SP + credentials + permissions)
# aad apps       — bulk listing of all app registrations with credential status
#
# Required permissions:
#   Application.Read.All (or Directory.Read.All)
#   for SP appRoleAssignments: AppRoleAssignment.ReadWrite.All or Directory.Read.All

from __future__ import annotations

import argparse
import json
from datetime import datetime, timedelta, timezone
from urllib.parse import quote

import httpx
from kiota_abstractions.base_request_configuration import RequestConfiguration
from loguru import logger
from rich.table import Table

from ...core.context import GraphContext
from ...utils import output, pagination
from ...utils.console import console
from ...utils.errors import handle_graph_errors

_BASE = "https://graph.microsoft.com/v1.0"
_NOW = None  # set at runtime


def _now() -> datetime:
    return datetime.now(tz=timezone.utc)


def _cred_status(end_dt: datetime | None) -> str:
    if end_dt is None:
        return "never-expires"
    now = _now()
    if end_dt < now:
        return "expired"
    if end_dt - now < timedelta(days=30):
        return "expiring-soon"
    return "valid"


def _status_badge(status: str) -> str:
    return {
        "never-expires": "[bold red]∞ never-expires[/bold red]",
        "expired": "[red]✗ expired[/red]",
        "expiring-soon": "[yellow]! expiring soon[/yellow]",
        "valid": "[green]✓ valid[/green]",
    }.get(status, status)


def _status_badge_short(status: str) -> str:
    return {
        "never-expires": "[bold red]∞[/bold red]",
        "expired": "[red]✗[/red]",
        "expiring-soon": "[yellow]![/yellow]",
        "valid": "[green]✓[/green]",
    }.get(status, "?")


# ---------------------------------------------------------------------------
# aad app <id> — single enriched view
# ---------------------------------------------------------------------------

def add_arguments_single(parser: argparse.ArgumentParser) -> None:
    parser.add_argument(
        "app_id",
        help="Application client ID (appId) or object ID.",
    )


async def _resolve_app(token: str, app_ref: str) -> dict | None:
    headers = {"Authorization": f"Bearer {token}"}
    async with httpx.AsyncClient(verify=False, timeout=30) as client:
        # Try object ID first
        r = await client.get(f"{_BASE}/applications/{quote(app_ref)}", headers=headers)
        if r.status_code == 200:
            return r.json()
        # Try appId filter
        r2 = await client.get(
            f"{_BASE}/applications?$filter=appId eq '{app_ref}'&$top=1",
            headers=headers,
        )
        if r2.status_code == 200:
            vals = r2.json().get("value", [])
            if vals:
                return vals[0]
    return None


async def _batch(token: str, requests: list[dict]) -> list[dict]:
    async with httpx.AsyncClient(verify=False, timeout=30) as client:
        resp = await client.post(
            f"{_BASE}/$batch",
            json={"requests": requests},
            headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
        )
    if resp.status_code != 200:
        raise RuntimeError(f"$batch failed: {resp.status_code} {resp.text[:200]}")
    return resp.json()["responses"]


@handle_graph_errors
async def run_single(context: GraphContext, args: argparse.Namespace) -> int:
    token = await context.get_access_token()
    if not token:
        logger.error("No access token available.")
        return 1

    logger.info(f"Resolving application: {args.app_id}")
    app = await _resolve_app(token, args.app_id)
    if not app:
        logger.error(f"Application not found: {args.app_id}")
        return 1

    obj_id = app["id"]
    app_client_id = app["appId"]

    # Batch: owners + SP lookup in one round-trip
    responses = await _batch(token, [
        {"id": "owners", "method": "GET", "url": f"/applications/{obj_id}/owners"},
        {"id": "sp", "method": "GET",
         "url": f"/servicePrincipals?$filter=appId eq '{app_client_id}'&$top=1"},
    ])

    app["owners"] = next(
        (r["body"].get("value", []) for r in responses if r["id"] == "owners" and r["status"] == 200),
        [],
    )

    sp_list = next(
        (r["body"].get("value", []) for r in responses if r["id"] == "sp" and r["status"] == 200),
        [],
    )
    sp = sp_list[0] if sp_list else None

    if sp:
        sp_id = sp["id"]
        sp_resp = await _batch(token, [
            {"id": "appRoles", "method": "GET",
             "url": f"/servicePrincipals/{sp_id}/appRoleAssignments"},
            {"id": "grants", "method": "GET",
             "url": f"/servicePrincipals/{sp_id}/oauth2PermissionGrants"},
        ])
        sp["appRoleAssignments"] = next(
            (r["body"].get("value", []) for r in sp_resp if r["id"] == "appRoles" and r["status"] == 200),
            [],
        )
        sp["oauth2PermissionGrants"] = next(
            (r["body"].get("value", []) for r in sp_resp if r["id"] == "grants" and r["status"] == 200),
            [],
        )
        app["servicePrincipal"] = sp

    if context.json_output:
        output.print_json(app)
        return 0

    _print_app(app)
    return 0


def _fmt_dt(dt_str: str | None) -> str:
    if not dt_str:
        return "[dim]—[/dim]"
    return dt_str[:19].replace("T", " ") + " UTC"


def _print_app(app: dict) -> None:
    console.print(f"\n[bold]App: {app.get('displayName', '?')}[/bold]")
    console.rule()
    console.print(f"  [bold]App ID (client)[/bold] : {app.get('appId')}")
    console.print(f"  [bold]Object ID[/bold]       : {app.get('id')}")
    console.print(f"  [bold]Audience[/bold]        : {app.get('signInAudience', '?')}")
    console.print(f"  [bold]Created[/bold]         : {_fmt_dt(app.get('createdDateTime'))}")

    # Owners
    owners = app.get("owners", [])
    if owners:
        names = ", ".join(o.get("displayName") or o.get("userPrincipalName") or o.get("id", "?") for o in owners)
        console.print(f"  [bold]Owners[/bold]          : {names}")

    # Password credentials
    pw_creds = app.get("passwordCredentials", [])
    console.print(f"\n[bold]Password Credentials[/bold] ({len(pw_creds)})")
    if pw_creds:
        for c in pw_creds:
            end = c.get("endDateTime")
            end_dt = datetime.fromisoformat(end.replace("Z", "+00:00")) if end else None
            status = _cred_status(end_dt)
            label = c.get("displayName") or c.get("hint") or str(c.get("keyId", "?"))
            end_str = end_dt.strftime("%Y-%m-%d %H:%M UTC") if end_dt else "never"
            console.print(f"  {_status_badge(status)}  {label}  expires: {end_str}")
    else:
        console.print("  [dim]none[/dim]")

    # Key credentials (certs)
    kc = app.get("keyCredentials", [])
    console.print(f"\n[bold]Key Credentials / Certs[/bold] ({len(kc)})")
    if kc:
        for c in kc:
            end = c.get("endDateTime")
            end_dt = datetime.fromisoformat(end.replace("Z", "+00:00")) if end else None
            status = _cred_status(end_dt)
            label = c.get("displayName") or str(c.get("keyId", "?"))
            ktype = c.get("type", "?")
            end_str = end_dt.strftime("%Y-%m-%d %H:%M UTC") if end_dt else "never"
            console.print(f"  {_status_badge(status)}  {label}  type: {ktype}  expires: {end_str}")
    else:
        console.print("  [dim]none[/dim]")

    # Service principal + permissions
    sp = app.get("servicePrincipal")
    if not sp:
        console.print("\n[dim]No service principal found in this tenant.[/dim]")
        return

    console.print(f"\n[bold]Service Principal[/bold]")
    console.print(f"  SP Object ID : {sp.get('id')}")
    console.print(f"  Type         : {sp.get('servicePrincipalType', '?')}")
    console.print(f"  Enabled      : {sp.get('accountEnabled', '?')}")

    roles = sp.get("appRoleAssignments", [])
    console.print(f"\n[bold]App Role Assignments (API permissions granted)[/bold] ({len(roles)})")
    if roles:
        for r in roles:
            console.print(
                f"  resource: {r.get('resourceDisplayName', '?')}  "
                f"role: {r.get('appRoleId', '?')}"
            )
    else:
        console.print("  [dim]none[/dim]")

    grants = sp.get("oauth2PermissionGrants", [])
    console.print(f"\n[bold]OAuth2 Delegated Grants[/bold] ({len(grants)})")
    if grants:
        for g in grants:
            console.print(
                f"  scope: {g.get('scope', '?')}  "
                f"resource: {g.get('resourceId', '?')}  "
                f"consent: {g.get('consentType', '?')}"
            )
    else:
        console.print("  [dim]none[/dim]")


# ---------------------------------------------------------------------------
# aad apps — bulk listing
# ---------------------------------------------------------------------------

def add_arguments_bulk(parser: argparse.ArgumentParser) -> None:
    parser.add_argument(
        "--with-secrets",
        action="store_true",
        help="Only apps that have at least one credential (secret or cert).",
    )
    parser.add_argument(
        "--expired",
        action="store_true",
        help="Only apps with at least one expired credential.",
    )
    parser.add_argument(
        "--expiring-soon",
        metavar="DAYS",
        type=int,
        default=None,
        help="Only apps with a credential expiring within N days.",
    )


@handle_graph_errors
async def run_bulk(context: GraphContext, args: argparse.Namespace) -> int:
    from msgraph.generated.applications.applications_request_builder import ApplicationsRequestBuilder

    logger.info("Enumerating app registrations...")

    QueryParams = ApplicationsRequestBuilder.ApplicationsRequestBuilderGetQueryParameters
    config = RequestConfiguration(
        query_parameters=QueryParams(
            select=["displayName", "appId", "id", "createdDateTime",
                    "signInAudience", "passwordCredentials", "keyCredentials"],
            top=999,
        )
    )
    apps_raw = await pagination.collect_all(
        context.graph_client.applications,
        request_configuration=config,
    )

    now = _now()

    def _annotate(creds: list, is_pw: bool) -> list[dict]:
        out = []
        for c in creds:
            end_dt = c.end_date_time
            status = _cred_status(end_dt)
            out.append({
                "keyId": str(c.key_id) if c.key_id else None,
                "displayName": c.display_name or (c.hint if is_pw else None),
                "endDateTime": end_dt.isoformat() if end_dt else None,
                "status": status,
            })
        return out

    apps = []
    for a in apps_raw:
        pw = _annotate(a.password_credentials or [], is_pw=True)
        kc = _annotate(a.key_credentials or [], is_pw=False)
        apps.append({
            "displayName": a.display_name,
            "appId": a.app_id,
            "id": a.id,
            "createdDateTime": a.created_date_time.isoformat() if a.created_date_time else None,
            "signInAudience": a.sign_in_audience,
            "_pw": pw,
            "_kc": kc,
        })

    # Apply filters
    if args.with_secrets:
        apps = [a for a in apps if a["_pw"] or a["_kc"]]
    if args.expired:
        apps = [a for a in apps if
                any(c["status"] == "expired" for c in a["_pw"] + a["_kc"])]
    if args.expiring_soon is not None:
        cutoff = now + timedelta(days=args.expiring_soon)
        apps = [a for a in apps if
                any(
                    c["endDateTime"] and
                    datetime.fromisoformat(c["endDateTime"]) <= cutoff
                    for c in a["_pw"] + a["_kc"]
                )]

    logger.info(f"{len(apps)} app(s) matched.")

    if context.json_output:
        output.print_json([{k: v for k, v in a.items() if not k.startswith("_")} |
                           {"passwordCredentials": a["_pw"], "keyCredentials": a["_kc"]}
                           for a in apps])
        return 0

    if not apps:
        console.print("[dim]No matching applications.[/dim]")
        return 0

    _print_bulk_table(apps)
    return 0


def _badges_for(creds: list[dict]) -> str:
    if not creds:
        return "[dim]-[/dim]"
    return " ".join(_status_badge_short(c["status"]) for c in creds)


def _print_bulk_table(apps: list[dict]) -> None:
    table = Table(show_header=True, header_style="bold", show_lines=False)
    table.add_column("Display Name", style="bold")
    table.add_column("App ID (client)", style="cyan", no_wrap=True)
    table.add_column("Secrets", justify="center")
    table.add_column("Certs", justify="center")
    table.add_column("Created", no_wrap=True)

    for app in sorted(apps, key=lambda a: (a.get("displayName") or "").lower()):
        created = (app.get("createdDateTime") or "")[:10]
        table.add_row(
            app.get("displayName") or "[dim]?[/dim]",
            app.get("appId") or "",
            _badges_for(app["_pw"]),
            _badges_for(app["_kc"]),
            created,
        )

    console.print(table)
    console.print(
        "[dim]Legend: [bold red]∞[/bold red] never-expires  "
        "[red]✗[/red] expired  [yellow]![/yellow] <30d  [green]✓[/green] valid[/dim]"
    )
