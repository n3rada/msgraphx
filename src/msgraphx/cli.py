# msrgraphx/cli.py

# Built-in imports
import argparse
import importlib
import json
import os
import pkgutil
import time
from pathlib import Path

# External library imports
import httpx
from azure.identity.aio import ClientSecretCredential
from loguru import logger

from msgraph import GraphServiceClient
from msgraph.generated.service_principals.service_principals_request_builder import (
    ServicePrincipalsRequestBuilder,
)
from msgraph.generated.models.o_data_errors.o_data_error import ODataError

# Local library imports
from . import __version__
from .core.context import GraphContext
from .modules import sharepoint, aad, me, outlook
from .utils import logbook, tokens


def load_subcommands_from_module(
    parent_parser: argparse.ArgumentParser,
    module_package,
    module_name: str,
    global_parent: argparse.ArgumentParser,
):
    """
    Dynamically load and register subcommands from a module package.

    This function discovers all modules within a package that contain
    'add_arguments' and 'run_with_arguments' functions, and registers them
    as subparsers.

    Args:
        parent_parser: The parent ArgumentParser to add subparsers to
        module_package: The package module to scan for subcommands
        module_name: The name to use for the main subcommand (e.g., 'sp', 'sharepoint')
        global_parent: Optional parent parser with global options to inherit

    Example:
        >>> sp_parser = subparsers.add_parser('sharepoint', aliases=['sp'])
        >>> load_subcommands_from_module(sp_parser, sharepoint, 'sharepoint', parent_parser)
    """
    subparsers = parent_parser.add_subparsers(
        dest=f"{module_name}_command", help=f"{module_name} subcommand"
    )

    for importer, full_module_name, is_pkg in pkgutil.iter_modules(
        module_package.__path__, module_package.__name__ + "."
    ):
        # Skip __init__ and __pycache__
        if full_module_name.endswith("__init__"):
            continue

        # Get the actual module name without the package prefix
        short_name = full_module_name.split(".")[-1]

        # Import the module
        try:
            module = importlib.import_module(full_module_name)
        except Exception:
            continue  # Skip modules that fail to load

        # Check if module has the required functions
        if not (
            hasattr(module, "add_arguments") and hasattr(module, "run_with_arguments")
        ):
            continue

        # Create subparser for this command, inheriting global options if available
        parents = [global_parent] if global_parent else []
        cmd_parser = subparsers.add_parser(
            short_name, help=f"{short_name.capitalize()} commands", parents=parents
        )
        module.add_arguments(cmd_parser)
        # Store the module reference for later execution
        cmd_parser.set_defaults(**{f"{module_name}_module": module})


def build_parser() -> argparse.ArgumentParser:
    # Create parent parser with global options that all subcommands inherit
    parent_parser = argparse.ArgumentParser(add_help=False)

    parent_parser.add_argument(
        "--debug",
        action="store_true",
        help="Enable debug logging (shortcut for --log-level DEBUG).",
    )

    parent_parser.add_argument(
        "--trace",
        action="store_true",
        help="Enable TRACE logging (shortcut for --log-level TRACE).",
    )

    parent_parser.add_argument(
        "--log-level",
        type=str,
        choices=["TRACE", "DEBUG", "INFO", "WARNING", "ERROR", "CRITICAL"],
        default=None,
        help="Set the logging level explicitly.",
    )

    parent_parser.add_argument(
        "--before",
        type=str,
        help="Filter items created on or before this date/time. Format: YYYY-MM-DD or relative (e.g. 5h, 3d, 1w, 2y).",
    )

    parent_parser.add_argument(
        "--after",
        type=str,
        help="Filter items created on or after this date/time. Format: YYYY-MM-DD or relative (e.g. 5h, 3d, 1w, 2y).",
    )

    parent_parser.add_argument(
        "--save",
        "--output",
        "-o",
        type=str,
        metavar="PATH",
        help="Directory path to save downloaded files. Creates the directory if it doesn't exist.",
    )

    parser = argparse.ArgumentParser(
        prog="msgraphx",
        add_help=True,
        description="Microsoft Graph eXploitation toolkit.",
        parents=[parent_parser],
    )

    parser.add_argument(
        "--version",
        action="version",
        version=f"%(prog)s {__version__}",
        help="Show version and exit.",
    )

    parser.add_argument(
        "--proxy",
        type=str,
        nargs="?",
        const="http://127.0.0.1:8080",
        help="Set HTTP(S) proxy, e.g. 'http://127.0.0.1:8080'. Default is Burp Suite proxy.",
    )

    # Authentication options
    auth_group = parser.add_argument_group(
        "Authentication",
        "Provide either token-based auth OR client credentials.",
    )

    # Mutually exclusive group for the primary auth method indicators
    auth_methods = auth_group.add_mutually_exclusive_group(required=False)

    auth_methods.add_argument(
        "--access-token",
        type=str,
        default=None,
        help="Microsoft Graph JWT token for delegated auth. If not provided, checks ACCESS_TOKEN env var, then .roadtools_auth file.",
    )

    auth_methods.add_argument(
        "--tenant-id",
        type=str,
        default=None,
        help="Azure AD Tenant ID for client credentials auth. If not provided, checks TENANT_ID env var.",
    )

    # Supporting arguments for token auth
    auth_group.add_argument(
        "--refresh-token",
        type=str,
        default=None,
        help="Microsoft Graph refresh token (used with --access-token). If not provided, checks REFRESH_TOKEN env var.",
    )

    # Supporting arguments for client credentials auth
    auth_group.add_argument(
        "--client-id",
        type=str,
        default=None,
        help="Azure AD Application (Client) ID (used with --tenant-id). If not provided, checks CLIENT_ID env var.",
    )

    auth_group.add_argument(
        "--client-secret",
        type=str,
        default=None,
        help="Azure AD Client Secret (used with --tenant-id). If not provided, checks CLIENT_SECRET env var.",
    )

    # Region for application permissions
    auth_group.add_argument(
        "--region",
        type=str,
        default="EMEA",
        help="Region for search requests with application permissions. Common values: NAM (North America), EMEA (Europe, Middle East, and Africa), APC (Asia Pacific). Defaults to EMEA.",
    )

    # Additional connection options
    parser.add_argument(
        "--drive-id",
        type=str,
        default=None,
        help="SharePoint Drive ID. If not provided, checks DRIVE_ID env var.",
    )

    subparsers = parser.add_subparsers(dest="command", help="Subcommand to run")

    # SharePoint subcommand (inherits parent_parser for global options)
    sp_parser = subparsers.add_parser(
        "sharepoint",
        aliases=["sp"],
        help="SharePoint commands",
        parents=[parent_parser],
    )
    load_subcommands_from_module(sp_parser, sharepoint, "sp", parent_parser)

    # Azure AD subcommand (inherits parent_parser for global options)
    aad_parser = subparsers.add_parser(
        "aad",
        aliases=["ad"],
        help="Azure Active Directory commands",
        parents=[parent_parser],
    )
    load_subcommands_from_module(aad_parser, aad, "aad", parent_parser)

    # Me subcommand (inherits parent_parser for global options)
    me_parser = subparsers.add_parser(
        "me",
        help="Current user information",
        parents=[parent_parser],
    )
    me.add_arguments(me_parser)

    # Outlook subcommand (inherits parent_parser for global options)
    outlook_parser = subparsers.add_parser(
        "outlook",
        aliases=["mail"],
        help="Outlook / mail commands",
        parents=[parent_parser],
    )
    load_subcommands_from_module(outlook_parser, outlook, "outlook", parent_parser)

    return parser


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _configure_logging(args) -> None:
    if args.log_level:
        level = args.log_level
    elif args.debug:
        level = "DEBUG"
    elif args.trace:
        level = "TRACE"
    else:
        level = "INFO"
    logbook.setup_logging(level=level)


def _apply_proxy(proxy: str | None) -> int:
    """Return 1 on bad format, 0 otherwise."""
    if not proxy:
        return 0
    if not proxy.startswith(("http://", "https://")):
        logger.error("Invalid proxy format.")
        return 1
    for key in ("http_proxy", "https_proxy", "HTTP_PROXY", "HTTPS_PROXY", "ALL_PROXY", "all_proxy"):
        os.environ[key] = proxy
    logger.info(f"🌐 Proxy set to {proxy}")
    return 0


def _check_public_ip() -> tuple[str | None, float]:
    try:
        with httpx.Client(verify=False, http1=True, http2=False) as client:
            t0 = time.perf_counter()
            r = client.get("https://api.ipify.org?format=json", timeout=15)
            rtt = time.perf_counter() - t0
            if r.status_code == 200:
                return r.json().get("ip"), rtt
    except httpx.TimeoutException:
        logger.error("Request timed-out.")
    except Exception as exc:
        logger.warning(f"⚠️ Error retrieving public IP: {exc}")
    return None, 0.0


def _load_token(args) -> tuple[str | None, str | None]:
    """Resolve access/refresh tokens from args → env → .roadtools_auth."""
    if args.access_token:
        return args.access_token, args.refresh_token

    env_token = os.environ.get("ACCESS_TOKEN")
    if env_token:
        logger.info("🔐 Using JWT from environment variable ACCESS_TOKEN.")
        return env_token, os.environ.get("REFRESH_TOKEN")

    logger.info("🔐 No JWT provided, checking for .roadtools_auth file.")
    tokens_path = Path(".roadtools_auth")
    if not tokens_path.exists():
        logger.error("🔐 No JWT provided and .roadtools_auth file not found.")
        return None, None

    logger.info("🔐 Using JWT from .roadtools_auth file.")
    try:
        data = json.loads(tokens_path.read_bytes())
        return data.get("accessToken"), data.get("refreshToken")
    except json.JSONDecodeError:
        logger.error("❌ Failed to decode .roadtools_auth file. Ensure it contains valid JSON.")
        return None, None


def _build_app_client(tenant_id: str, client_id: str, client_secret: str) -> "GraphServiceClient | None":
    try:
        credentials = ClientSecretCredential(
            tenant_id=tenant_id,
            client_id=client_id,
            client_secret=client_secret,
        )
        return GraphServiceClient(
            credentials=credentials,
            scopes=["https://graph.microsoft.com/.default"],
        )
    except Exception as exc:
        logger.error(f"❌ Failed to create Graph client with client credentials: {exc}")
        return None


def _classify_token(token) -> bool:
    """Return True if the token carries app-only permissions."""
    payload = token.payload
    if payload.get("roles") and not payload.get("scp"):
        logger.info("📋 Token contains application permissions (app roles).")
        return True
    logger.info("👤 Token contains delegated permissions (user scopes).")
    return False


async def _authenticate(args) -> tuple["GraphServiceClient | None", bool, int]:
    """Resolve auth method and return (client, is_app_only, exit_code)."""
    if args.access_token and args.tenant_id:
        logger.error(
            "❌ Cannot use both token-based authentication (--access-token) and "
            "client credentials (--tenant-id) simultaneously."
        )
        return None, False, 1

    has_env_creds = all(os.environ.get(k) for k in ("TENANT_ID", "CLIENT_ID", "CLIENT_SECRET"))
    use_creds = (
        args.tenant_id is not None
        or (has_env_creds and args.access_token is None and not os.environ.get("ACCESS_TOKEN"))
    )

    if use_creds:
        tenant_id = args.tenant_id or os.environ.get("TENANT_ID")
        client_id = args.client_id or os.environ.get("CLIENT_ID")
        client_secret = args.client_secret or os.environ.get("CLIENT_SECRET")
        args.tenant_id, args.client_id, args.client_secret = tenant_id, client_id, client_secret

        if tenant_id:
            logger.info(f"🏢 Tenant ID: {tenant_id}")
        if client_id:
            logger.info(f"📱 Client ID: {client_id}")

        logger.info("🔑 Using client credentials authentication")
        client = _build_app_client(tenant_id, client_id, client_secret)
        return (client, True, 0) if client else (None, True, 1)

    access_token, refresh_token = _load_token(args)
    if not access_token:
        return None, False, 1

    token = tokens.TokenManager(access_token, refresh_token)

    if token.is_expired:
        logger.error("🔒 Token is expired.")
        return None, False, 1

    if (
        token.audience != "00000003-0000-0000-c000-000000000000"
        and not token.audience.startswith("https://graph.microsoft.com")
    ):
        logger.error(f"❌ JWT audience mismatch, got '{token.audience}'")
        return None, False, 1

    is_app_only = _classify_token(token)
    token.start_auto_refresh()
    return GraphServiceClient(token), is_app_only, 0


async def _verify_connection(graph_client, is_app_only: bool, client_id: str | None, region: str):
    """Verify the Graph connection. Returns (cached_user, exit_code)."""
    if is_app_only:
        logger.info("✅ Connected to Microsoft Graph with application permissions.")
        logger.info(f"🌍 Region: {region}")
        if client_id:
            await _log_service_principal(graph_client, client_id)
        return None, 0

    try:
        user = await graph_client.me.get()
        logger.info(f"🔗 Connected to Microsoft Graph as: {user.display_name}")
        return user, 0
    except ODataError as exc:
        code = getattr(exc.error, "code", "Unknown")
        message = getattr(exc.error, "message", "No message")
        logger.error(f"❌ Failed to connect to Microsoft Graph API. Code: {code} | Message: {message}")
        if code == "InvalidAuthenticationToken":
            logger.error("🔒 Token is invalid or expired. Please re-authenticate.")
        return None, 1
    except TimeoutError:
        logger.error("❌ Connection timeout. Check your network connection.")
        return None, 1
    except Exception as exc:
        logger.error(f"❌ Unexpected error while connecting to Graph API: {exc}")
        return None, 1


async def _log_service_principal(graph_client, client_id: str) -> None:
    try:
        query_params = ServicePrincipalsRequestBuilder.ServicePrincipalsRequestBuilderGetQueryParameters(
            filter=f"appId eq '{client_id}'", select=["displayName", "appId"]
        )
        request_config = ServicePrincipalsRequestBuilder.ServicePrincipalsRequestBuilderGetRequestConfiguration(
            query_parameters=query_params
        )
        sps = await graph_client.service_principals.get(request_configuration=request_config)
        if sps and sps.value:
            logger.info(f"🤖 Application: {sps.value[0].display_name}")
    except Exception:
        pass


async def _dispatch(args, context) -> int:
    """Route to the appropriate subcommand handler."""
    command = args.command

    if command in ("sharepoint", "sp"):
        if not getattr(args, "sp_command", None):
            logger.error("Please specify a SharePoint subcommand (e.g., 'msgraphx sp search')")
            return 1
        return await args.sp_module.run_with_arguments(context, args)

    if command in ("aad", "ad"):
        if not getattr(args, "aad_command", None):
            logger.error("Please specify an Azure AD subcommand (e.g., 'msgraphx aad search admin')")
            return 1
        return await args.aad_module.run_with_arguments(context, args)

    if command == "me":
        return await me.run_with_arguments(context, args)

    if command in ("outlook", "mail"):
        if not getattr(args, "outlook_command", None):
            logger.error("Please specify an Outlook subcommand (e.g., 'msgraphx outlook contacts')")
            return 1
        return await args.outlook_module.run_with_arguments(context, args)

    logger.info("✅ Authentication successful. Use a subcommand to perform actions.")
    return 0


# ---------------------------------------------------------------------------
# Entry points
# ---------------------------------------------------------------------------

@logger.catch
async def _main() -> int:
    parser = build_parser()
    args = parser.parse_args()

    _configure_logging(args)

    if _apply_proxy(getattr(args, "proxy", None)):
        return 1

    public_ip, rtt = _check_public_ip()
    if not public_ip:
        return 1

    logger.info(f"🌍 Public IP: {public_ip} (⏱️ RTT: {rtt:.2f}s)")

    args.drive_id = getattr(args, "drive_id", None) or os.environ.get("DRIVE_ID")
    if args.drive_id:
        logger.info(f"💾 Drive ID: {args.drive_id}")

    graph_client, is_app_only, err = await _authenticate(args)
    if err:
        return err

    args.is_app_only = is_app_only

    cached_user, err = await _verify_connection(
        graph_client, is_app_only, getattr(args, "client_id", None), args.region
    )
    if err:
        return err

    context = GraphContext(
        graph_client=graph_client,
        is_app_only=is_app_only,
        region=args.region,
        cached_user=cached_user,
    )

    return await _dispatch(args, context)


def main() -> None:
    import asyncio

    raise SystemExit(asyncio.run(_main()))

