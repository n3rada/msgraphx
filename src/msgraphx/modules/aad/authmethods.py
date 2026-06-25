# msgraphx/modules/aad/authmethods.py
#
# Enumerate authentication methods registered for a user.
# Surfaces every MFA factor: phone/SMS, Authenticator app, FIDO2/passkey,
# software TOTP, Temporary Access Pass, Windows Hello, email OTP, password.

# Built-in imports
from __future__ import annotations

import argparse

# External library imports
from loguru import logger
from rich.table import Table

# Local library imports
from ...core.context import GraphContext
from ...utils import output
from ...utils.console import console
from ...utils.errors import handle_graph_errors
from ...utils.roles import require_scopes

_METHOD_LABEL: dict[str, str] = {
    "#microsoft.graph.phoneAuthenticationMethod": "Phone / SMS",
    "#microsoft.graph.microsoftAuthenticatorAuthenticationMethod": "Authenticator app",
    "#microsoft.graph.fido2AuthenticationMethod": "FIDO2 / Passkey",
    "#microsoft.graph.softwareOathAuthenticationMethod": "TOTP (software OATH)",
    "#microsoft.graph.temporaryAccessPassAuthenticationMethod": "Temporary Access Pass",
    "#microsoft.graph.windowsHelloForBusinessAuthenticationMethod": "Windows Hello",
    "#microsoft.graph.emailAuthenticationMethod": "Email OTP",
    "#microsoft.graph.passwordAuthenticationMethod": "Password",
}

def add_arguments(parser: "argparse.ArgumentParser") -> None:
    parser.add_argument(
        "user",
        help="User UPN or object ID (e.g. alice@corp.com or a GUID).",
    )

@handle_graph_errors
@require_scopes("UserAuthenticationMethod.Read.All")
async def run_with_arguments(
    context: "GraphContext", args: "argparse.Namespace"
) -> int:
    if context.is_app_only:
        logger.error("This module requires delegated authentication (user context).")
        return 1

    logger.info(f"Fetching auth methods for: {args.user}")

    result = await context.graph_client.users.by_user_id(
        args.user
    ).authentication.methods.get()

    methods = (result.value or []) if result else []

    if not methods:
        logger.info("No authentication methods found.")
        if context.json_output:
            output.print_json([])
        return 0

    rows = []
    for m in methods:
        kind = _METHOD_LABEL.get(m.odata_type or "", m.odata_type or "unknown")

        detail = ""
        t = m.odata_type or ""
        if t == "#microsoft.graph.phoneAuthenticationMethod":
            detail = getattr(m, "phone_number", "") or ""
            phone_type = getattr(m, "phone_type", None)
            if phone_type:
                detail += f" ({str(phone_type).split('.')[-1]})"
        elif t == "#microsoft.graph.microsoftAuthenticatorAuthenticationMethod":
            detail = getattr(m, "display_name", "") or ""
            version = getattr(m, "phone_app_version", None)
            if version:
                detail += f" v{version}"
        elif t == "#microsoft.graph.fido2AuthenticationMethod":
            detail = getattr(m, "display_name", "") or ""
            model_name = getattr(m, "model", None)
            if model_name:
                detail += f" ({model_name})"
        elif t == "#microsoft.graph.emailAuthenticationMethod":
            detail = getattr(m, "email_address", "") or ""

        created = ""
        if m.created_date_time:
            created = m.created_date_time.strftime("%Y-%m-%d")

        rows.append({
            "method_id": m.id,
            "type": t,
            "label": kind,
            "detail": detail,
            "registered": created,
        })

    if context.json_output:
        output.print_json(rows)
        return 0

    if context.ndjson_output:
        for row in rows:
            output.print_ndjson_item(row)
        return 0

    table = Table(show_header=True, header_style="bold", box=None, padding=(0, 1))
    table.add_column("#", style="dim", justify="right", width=4)
    table.add_column("Method", min_width=28)
    table.add_column("Detail", style="dim", min_width=30)
    table.add_column("Registered", style="cyan", width=12)

    for i, row in enumerate(rows, 1):
        table.add_row(str(i), row["label"], row["detail"], row["registered"])

    console.print(f"[bold]Auth methods: {args.user}[/bold]")
    console.rule()
    console.print(table)
    logger.success(f"{len(rows)} method(s) found.")
    return 0
