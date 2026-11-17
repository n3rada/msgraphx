# msgraphx/modules/aad/apps.py
#
# aad app  <id>  — enriched single-app profile (registration + SP + credentials + permissions)
# aad apps       — bulk listing of all app registrations with credential status

from __future__ import annotations

# Built-in imports
import argparse
import asyncio
from collections.abc import Mapping, Sequence
from datetime import datetime, timedelta, timezone
from typing import Any, TypeVar, cast

# External library imports
from kiota_abstractions.base_request_configuration import RequestConfiguration
from loguru import logger
from msgraph.generated.applications.applications_request_builder import ApplicationsRequestBuilder
from msgraph.generated.models.app_role_assignment import AppRoleAssignment
from msgraph.generated.models.application import Application
from msgraph.generated.models.directory_object import DirectoryObject
from msgraph.generated.models.o_auth2_permission_grant import OAuth2PermissionGrant
from msgraph.generated.models.service_principal import ServicePrincipal
from msgraph.generated.service_principals.service_principals_request_builder import (
    ServicePrincipalsRequestBuilder,
)
from rich.table import Table

# Local library imports
from ...core.context import GraphContext
from ...utils import odata, output, pagination
from ...utils.console import console
from ...utils.errors import handle_graph_errors, raise_if_forbidden
from ...utils.roles import require_scopes

_T = TypeVar("_T")


def _now() -> datetime:
    return datetime.now(tz=timezone.utc)

def _cred_status(end_dt: datetime | None) -> str:
    """Classify a credential by its expiry time.

    Args:
        end_dt: The credential's `endDateTime`. `None` means it never expires.

    Returns:
        One of `"never-expires"`, `"expired"`, `"expiring-soon"` (within 30
        days), or `"valid"`.
    """
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

def _owner_name(obj: DirectoryObject) -> str:
    """Return a display label for an app owner.

    Args:
        obj: Owner returned by the `owners` navigation. Its display name may
            live on the typed model or in `additional_data`.

    Returns:
        The display name, else the UPN, else the object ID, else `"?"`.
    """
    if hasattr(obj, "display_name") and obj.display_name:
        return obj.display_name
    ad = getattr(obj, "additional_data", {}) or {}
    return ad.get("displayName") or ad.get("userPrincipalName") or (obj.id or "?")

# ---------------------------------------------------------------------------
# aad app <id> — single enriched view
# ---------------------------------------------------------------------------

def add_arguments_single(parser: argparse.ArgumentParser) -> None:
    """Register the arguments for `aad app`.

    Args:
        parser: Parser for the `app` subcommand.
    """
    parser.add_argument(
        "app_id",
        help="Application client ID (appId) or object ID.",
    )

async def _resolve_app(context: GraphContext, app_ref: str) -> Application | None:
    """Look up an app registration by object ID, then by client ID.

    Args:
        context: Shared Graph runtime context.
        app_ref: Application object ID or client ID (`appId`).

    Returns:
        The matching application, or `None` if neither lookup finds one.

    Raises:
        ForbiddenGraphError: The token cannot read applications.
    """
    # Try object ID directly
    try:
        app = await context.graph_client.applications.by_application_id(app_ref).get()
        if app and app.id:
            return app
    except Exception as exc:
        raise_if_forbidden(exc)

    # Fall back to appId (client ID) filter
    try:
        QueryParams = ApplicationsRequestBuilder.ApplicationsRequestBuilderGetQueryParameters
        config = RequestConfiguration(
            query_parameters=QueryParams(filter=f"appId eq {odata.quote(app_ref)}", top=1)
        )
        resp = await context.graph_client.applications.get(request_configuration=config)
        if resp and resp.value:
            return resp.value[0]
    except Exception as exc:
        raise_if_forbidden(exc)
        raise

    return None

@handle_graph_errors
@require_scopes("Application.Read.All")
async def run_single(context: GraphContext, args: argparse.Namespace) -> int:
    """Show one app registration with its owners, credentials, and permissions.

    Also looks up the app's service principal in this tenant and its app role
    assignments and OAuth2 delegated grants. Those follow-up lookups degrade
    to empty sections with a warning when they fail, including on 403.
    Prints JSON under `--json`, otherwise a Rich report. Warns if the token
    lacks `Application.Read.All`.

    Args:
        context: Shared Graph runtime context.
        args: Parsed arguments; `args.app_id` is an object ID or client ID.

    Returns:
        0 on success, 1 if the application is not found.

    Raises:
        AuthenticationError: The Graph token is invalid or expired.
        ForbiddenGraphError: The token cannot read the application.
    """
    logger.info(f"Resolving application: {args.app_id}")
    app = await _resolve_app(context, args.app_id)
    if not app:
        logger.error(f"Application not found: {args.app_id}")
        return 1

    # Graph always returns `id` for an application fetched without $select.
    obj_id = cast(str, app.id)
    app_client_id = app.app_id

    # Owners + SP lookup in parallel
    SpQueryParams = ServicePrincipalsRequestBuilder.ServicePrincipalsRequestBuilderGetQueryParameters
    sp_config = RequestConfiguration(
        query_parameters=SpQueryParams(filter=f"appId eq {odata.quote(app_client_id or '')}", top=1)
    )

    owners_result, sp_result = await asyncio.gather(
        _safe_collect(context.graph_client.applications.by_application_id(obj_id).owners),
        _safe_get(context.graph_client.service_principals, sp_config),
        return_exceptions=True,
    )

    owners: list[DirectoryObject] = _settled(owners_result, "app owners") or []
    sp_resp = _settled(sp_result, "service principal")
    sp_items = (sp_resp.value if sp_resp and sp_resp.value else [])
    sp = sp_items[0] if sp_items else None

    role_assignments: list[AppRoleAssignment] = []
    grants: list[OAuth2PermissionGrant] = []
    if sp:
        sp_builder = context.graph_client.service_principals.by_service_principal_id(sp.id)
        roles_result: list[AppRoleAssignment] | BaseException
        grants_result: list[OAuth2PermissionGrant] | BaseException
        roles_result, grants_result = await asyncio.gather(
            _safe_collect(sp_builder.app_role_assignments),
            _safe_collect(sp_builder.oauth2_permission_grants),
            return_exceptions=True,
        )
        role_assignments = _settled(roles_result, "app role assignments") or []
        grants = _settled(grants_result, "OAuth2 permission grants") or []

    if context.json_output:
        output.print_json(_app_to_dict(app, owners, sp, role_assignments, grants))
        return 0

    _print_app(app, owners, sp, role_assignments, grants)
    return 0

def _settled(result: _T | BaseException, what: str) -> _T | None:
    """Unwrap one `asyncio.gather(..., return_exceptions=True)` result.

    A failed lookup (including a 403 surfaced as `ForbiddenGraphError`) is
    logged as a warning so the rest of the report can still be printed.

    Args:
        result: The gathered value or the exception it raised.
        what: Human-readable name of the lookup, used in the warning.

    Returns:
        The value, or `None` if the lookup raised an `Exception`.

    Raises:
        BaseException: The lookup raised a non-`Exception` such as
            cancellation, which is re-raised.
    """
    if isinstance(result, Exception):
        logger.warning(f"Could not fetch {what}: {result}")
        return None
    if isinstance(result, BaseException):
        raise result
    return result

async def _safe_collect(builder: Any) -> list[Any]:
    """Collect every page of a collection, degrading to empty on error.

    Args:
        builder: SDK request builder for the collection.

    Returns:
        All items, or an empty list if a non-403 error occurs.

    Raises:
        ForbiddenGraphError: The token lacks permission for the collection.
    """
    try:
        return await pagination.collect_all(builder)
    except Exception as exc:
        raise_if_forbidden(exc)
        logger.warning(f"Could not fetch collection: {exc}")
        return []

async def _safe_get(builder: Any, config: RequestConfiguration[Any] | None) -> Any:
    """Issue a single GET, degrading to `None` on error.

    Args:
        builder: SDK request builder to call `get` on.
        config: Request configuration for the GET.

    Returns:
        The builder's response, or `None` if a non-403 error occurs.

    Raises:
        ForbiddenGraphError: The token lacks permission for the resource.
    """
    try:
        return await builder.get(request_configuration=config)
    except Exception as exc:
        raise_if_forbidden(exc)
        logger.warning(f"Could not fetch resource: {exc}")
        return None

def _fmt_dt(dt: datetime | None) -> str:
    if not dt:
        return "[dim]—[/dim]"
    return dt.strftime("%Y-%m-%d %H:%M") + " UTC"

def _app_to_dict(
    app: Application,
    owners: Sequence[DirectoryObject],
    sp: ServicePrincipal | None,
    role_assignments: Sequence[AppRoleAssignment],
    grants: Sequence[OAuth2PermissionGrant],
) -> dict[str, Any]:
    """Build the `--json` representation of a single app.

    Args:
        app: The app registration.
        owners: The app's owners.
        sp: The app's service principal in this tenant, if any.
        role_assignments: App role assignments granted to `sp`.
        grants: OAuth2 delegated permission grants held by `sp`.

    Returns:
        A camelCase dict with the app's properties, credential status, owners
        and, when `sp` is given, a `servicePrincipal` entry.
    """
    pw = [
        {
            "keyId": str(c.key_id) if c.key_id else None,
            "displayName": c.display_name or c.hint,
            "endDateTime": c.end_date_time.isoformat() if c.end_date_time else None,
            "status": _cred_status(c.end_date_time),
        }
        for c in (app.password_credentials or [])
    ]
    kc = [
        {
            "keyId": str(c.key_id) if c.key_id else None,
            "displayName": c.display_name,
            "endDateTime": c.end_date_time.isoformat() if c.end_date_time else None,
            "type": c.type,
            "status": _cred_status(c.end_date_time),
        }
        for c in (app.key_credentials or [])
    ]
    result: dict[str, Any] = {
        "displayName": app.display_name,
        "appId": app.app_id,
        "id": app.id,
        "signInAudience": app.sign_in_audience,
        "createdDateTime": app.created_date_time.isoformat() if app.created_date_time else None,
        "owners": [_owner_name(o) for o in owners],
        "passwordCredentials": pw,
        "keyCredentials": kc,
    }
    if sp:
        result["servicePrincipal"] = {
            "id": sp.id,
            "servicePrincipalType": sp.service_principal_type,
            "accountEnabled": sp.account_enabled,
            "appRoleAssignments": [
                {
                    "resourceDisplayName": r.resource_display_name,
                    "appRoleId": str(r.app_role_id) if r.app_role_id else None,
                }
                for r in role_assignments
            ],
            "oauth2PermissionGrants": [
                {
                    "scope": g.scope,
                    "resourceId": g.resource_id,
                    "consentType": g.consent_type,
                }
                for g in grants
            ],
        }
    return result

def _print_app(
    app: Application,
    owners: Sequence[DirectoryObject],
    sp: ServicePrincipal | None,
    role_assignments: Sequence[AppRoleAssignment],
    grants: Sequence[OAuth2PermissionGrant],
) -> None:
    """Print a single app as a Rich report.

    Args:
        app: The app registration.
        owners: The app's owners.
        sp: The app's service principal in this tenant, if any.
        role_assignments: App role assignments granted to `sp`.
        grants: OAuth2 delegated permission grants held by `sp`.
    """
    console.print(f"\n[bold]App: {app.display_name or '?'}[/bold]")
    console.rule()
    console.print(f"  [bold]App ID (client)[/bold] : {app.app_id}")
    console.print(f"  [bold]Object ID[/bold]       : {app.id}")
    console.print(f"  [bold]Audience[/bold]        : {app.sign_in_audience or '?'}")
    console.print(f"  [bold]Created[/bold]         : {_fmt_dt(app.created_date_time)}")

    if owners:
        console.print(f"  [bold]Owners[/bold]          : {', '.join(_owner_name(o) for o in owners)}")

    pw_creds = app.password_credentials or []
    console.print(f"\n[bold]Password Credentials[/bold] ({len(pw_creds)})")
    if pw_creds:
        for c in pw_creds:
            status = _cred_status(c.end_date_time)
            label = c.display_name or c.hint or str(c.key_id or "?")
            end_str = _fmt_dt(c.end_date_time) if c.end_date_time else "never"
            console.print(f"  {_status_badge(status)}  {label}  expires: {end_str}")
    else:
        console.print("  [dim]none[/dim]")

    kc = app.key_credentials or []
    console.print(f"\n[bold]Key Credentials / Certs[/bold] ({len(kc)})")
    if kc:
        for cert in kc:
            status = _cred_status(cert.end_date_time)
            label = cert.display_name or str(cert.key_id or "?")
            end_str = _fmt_dt(cert.end_date_time) if cert.end_date_time else "never"
            console.print(f"  {_status_badge(status)}  {label}  type: {cert.type}  expires: {end_str}")
    else:
        console.print("  [dim]none[/dim]")

    if not sp:
        console.print("\n[dim]No service principal found in this tenant.[/dim]")
        return

    console.print("\n[bold]Service Principal[/bold]")
    console.print(f"  SP Object ID : {sp.id}")
    console.print(f"  Type         : {sp.service_principal_type or '?'}")
    console.print(f"  Enabled      : {sp.account_enabled}")

    console.print(
        f"\n[bold]App Role Assignments (API permissions granted)[/bold] ({len(role_assignments)})"
    )
    if role_assignments:
        for r in role_assignments:
            console.print(f"  resource: {r.resource_display_name or '?'}  role: {r.app_role_id}")
    else:
        console.print("  [dim]none[/dim]")

    console.print(f"\n[bold]OAuth2 Delegated Grants[/bold] ({len(grants)})")
    if grants:
        for g in grants:
            console.print(
                f"  scope: {g.scope or '?'}  resource: {g.resource_id or '?'}  consent: {g.consent_type or '?'}"
            )
    else:
        console.print("  [dim]none[/dim]")

# ---------------------------------------------------------------------------
# aad apps — bulk listing
# ---------------------------------------------------------------------------

def add_arguments_bulk(parser: argparse.ArgumentParser) -> None:
    """Register the filter arguments for `aad apps`.

    Args:
        parser: Parser for the `apps` subcommand.
    """
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
@require_scopes("Application.Read.All")
async def run_bulk(context: GraphContext, args: argparse.Namespace) -> int:
    """List every app registration with the status of its credentials.

    Prints JSON under `--json`, otherwise a Rich table. Warns if the token
    lacks `Application.Read.All`.

    Args:
        context: Shared Graph runtime context.
        args: Parsed arguments. `args.with_secrets`, `args.expired`, and
            `args.expiring_soon` (days, or `None`) filter the apps.

    Returns:
        0 on success.

    Raises:
        AuthenticationError: The Graph token is invalid or expired.
        ForbiddenGraphError: The token cannot list applications.
    """
    logger.info("Enumerating app registrations...")

    QueryParams = ApplicationsRequestBuilder.ApplicationsRequestBuilderGetQueryParameters
    config = RequestConfiguration(
        query_parameters=QueryParams(
            select=["displayName", "appId", "id", "createdDateTime",
                    "signInAudience", "passwordCredentials", "keyCredentials"],
            top=999,
        )
    )
    try:
        apps_raw = await pagination.collect_all(
            context.graph_client.applications,
            request_configuration=config,
        )
    except Exception as exc:
        raise_if_forbidden(exc)
        raise

    now = _now()

    # `creds` holds PasswordCredential items when is_pw, else KeyCredential;
    # only the former has `hint`, which the type checker cannot see.
    def _annotate_creds(creds: Sequence[Any], is_pw: bool) -> list[dict[str, Any]]:
        out = []
        for c in creds:
            status = _cred_status(c.end_date_time)
            out.append({
                "keyId": str(c.key_id) if c.key_id else None,
                "displayName": (c.display_name or c.hint) if is_pw else c.display_name,
                "endDateTime": c.end_date_time.isoformat() if c.end_date_time else None,
                "status": status,
            })
        return out

    apps = []
    for a in apps_raw:
        pw = _annotate_creds(a.password_credentials or [], is_pw=True)
        kc = _annotate_creds(a.key_credentials or [], is_pw=False)
        apps.append({
            "displayName": a.display_name,
            "appId": a.app_id,
            "id": a.id,
            "createdDateTime": a.created_date_time.isoformat() if a.created_date_time else None,
            "signInAudience": a.sign_in_audience,
            "_pw": pw,
            "_kc": kc,
        })

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
        output.print_json([
            {k: v for k, v in a.items() if not k.startswith("_")} |
            {"passwordCredentials": a["_pw"], "keyCredentials": a["_kc"]}
            for a in apps
        ])
        return 0

    if not apps:
        console.print("[dim]No matching applications.[/dim]")
        return 0

    _print_bulk_table(apps)
    return 0

def _badges_for(creds: Sequence[Mapping[str, Any]]) -> str:
    if not creds:
        return "[dim]-[/dim]"
    return " ".join(_status_badge_short(c["status"]) for c in creds)

def _print_bulk_table(apps: Sequence[Mapping[str, Any]]) -> None:
    """Print the app list as a Rich table sorted by display name.

    Args:
        apps: App rows built by `run_bulk`, including the private `_pw` and
            `_kc` credential lists.
    """
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
