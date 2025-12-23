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
from .modules import sharepoint, aad, me
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

    return parser


@logger.catch
async def main() -> int:

    parser = build_parser()
    args = parser.parse_args()

    # Determine log level: --log-level takes precedence, then --debug, then --trace, then default INFO
    if args.log_level:
        log_level = args.log_level
    elif args.debug:
        log_level = "DEBUG"
    elif args.trace:
        log_level = "TRACE"
    else:
        log_level = "INFO"

    logbook.setup_logging(level=log_level)

    if args.proxy:
        if not args.proxy.startswith(("http://", "https://")):
            logger.error("Invalid proxy format.")
            return 1

        for k in [
            "http_proxy",
            "https_proxy",
            "HTTP_PROXY",
            "HTTPS_PROXY",
            "ALL_PROXY",
            "all_proxy",
        ]:
            os.environ[k] = args.proxy

        logger.info(f"🌐 Proxy set to {args.proxy}")

    public_ip = None
    try:
        with httpx.Client(verify=False, http1=True, http2=False) as client:
            start_time = time.perf_counter()
            response = client.get("https://api.ipify.org?format=json", timeout=15)
            end_time = time.perf_counter()

            rtt = end_time - start_time

            if response.status_code == 200:
                public_ip = response.json().get("ip", None)

    except httpx.TimeoutException:
        logger.error("Request timed-out.")
    except Exception as e:
        logger.warning(f"⚠️ Error retrieving public IP: {e}")

    if public_ip is None:
        return 1

    logger.info(f"🌍 Public IP: {public_ip} (⏱️ RTT: {rtt:.2f}s)")

    # Check for explicit authentication method conflicts (command-line args only)
    if args.access_token and args.tenant_id:
        logger.error(
            "❌ Cannot use both token-based authentication (--access-token) and "
            "client credentials (--tenant-id, --client-id, --client-secret) simultaneously."
        )
        return 1

    # Determine authentication method with priority: explicit args > env vars
    # Token auth takes precedence over client credentials if both are in environment
    has_explicit_token = args.access_token is not None
    has_explicit_client_creds = args.tenant_id is not None
    has_env_token = os.environ.get("ACCESS_TOKEN") is not None
    has_env_client_creds = (
        os.environ.get("TENANT_ID")
        and os.environ.get("CLIENT_ID")
        and os.environ.get("CLIENT_SECRET")
    )

    # Prioritize: explicit token > explicit client creds > env token > env client creds
    use_token_auth = has_explicit_token or (
        has_env_token and not has_explicit_client_creds
    )
    use_client_creds = not use_token_auth and (
        has_explicit_client_creds or has_env_client_creds
    )

    # Load credentials based on chosen method
    tenant_id = None
    client_id = None
    client_secret = None

    if use_client_creds:
        tenant_id = args.tenant_id or os.environ.get("TENANT_ID")
        client_id = args.client_id or os.environ.get("CLIENT_ID")
        client_secret = args.client_secret or os.environ.get("CLIENT_SECRET")

    # Always load drive_id regardless of auth method
    drive_id = args.drive_id or os.environ.get("DRIVE_ID")

    # Store in args for modules to access
    args.tenant_id = tenant_id
    args.client_id = client_id
    args.client_secret = client_secret
    args.drive_id = drive_id
    args.is_app_only = False  # Will be set later after auth type is determined

    if tenant_id:
        logger.info(f"🏢 Tenant ID: {tenant_id}")
    if client_id:
        logger.info(f"📱 Client ID: {client_id}")
    if drive_id:
        logger.info(f"💾 Drive ID: {drive_id}")

    # Build Graph client based on authentication method
    graph_client = None
    is_app_only = False  # Track if using application-only permissions

    if use_client_creds and tenant_id and client_id and client_secret:
        # Use client credentials flow - always app-only
        logger.info("🔑 Using client credentials authentication")
        is_app_only = True
        try:
            credentials = ClientSecretCredential(
                tenant_id=tenant_id,
                client_id=client_id,
                client_secret=client_secret,
            )
            scopes = ["https://graph.microsoft.com/.default"]
            graph_client = GraphServiceClient(credentials=credentials, scopes=scopes)
        except Exception as e:
            logger.error(
                f"❌ Failed to create Graph client with client credentials: {e}"
            )
            return 1
    else:
        # Use token-based authentication
        access_token = args.access_token
        refresh_token = args.refresh_token

        if not access_token:
            # Try environment variable
            access_token = os.environ.get("ACCESS_TOKEN")
            refresh_token = os.environ.get("REFRESH_TOKEN")

            if access_token:
                logger.info("🔐 Using JWT from environment variable ACCESS_TOKEN.")
            else:
                # Fall back to .roadtools_auth file
                logger.info("🔐 No JWT provided, checking for .roadtools_auth file.")
                tokens_path = Path(".roadtools_auth")
                if not tokens_path.exists():
                    logger.error(
                        "🔐 No JWT provided and .roadtools_auth file not found."
                    )
                    return 1

                logger.info("🔐 Using JWT from .roadtools_auth file.")
                try:
                    data = json.loads(tokens_path.read_bytes())
                    access_token = data.get("accessToken")
                    refresh_token = data.get("refreshToken")
                except json.JSONDecodeError:
                    logger.error(
                        "❌ Failed to decode .roadtools_auth file. Ensure it contains valid JSON."
                    )
                    return 1

        token = tokens.TokenManager(access_token, refresh_token)

        if token.is_expired:
            logger.error("🔒 Token is expired.")
            return 1

        if not (
            token.audience == "00000003-0000-0000-c000-000000000000"
            or token.audience.startswith("https://graph.microsoft.com")
        ):
            logger.error(f"❌ JWT audience mismatch, got '{token.audience}'")
            return 1

        # Check if token has application permissions (roles) or delegated permissions (scp)
        token_payload = token.payload
        has_roles = "roles" in token_payload and token_payload["roles"]
        has_scp = "scp" in token_payload and token_payload["scp"]

        if has_roles and not has_scp:
            is_app_only = True
            logger.info("📋 Token contains application permissions (app roles).")
        elif has_scp:
            logger.info("👤 Token contains delegated permissions (user scopes).")

        token.start_auto_refresh()

        # Build Graph client with token
        graph_client = GraphServiceClient(token)

    # Store authentication type for modules to access
    args.is_app_only = is_app_only

    # Verify connection - only check /me endpoint for delegated auth
    cached_user = None
    if is_app_only:
        # For app-only permissions, we can't use /me endpoint (no signed-in user)
        # But we can verify by getting the service principal info
        logger.info("✅ Connected to Microsoft Graph with application permissions.")
        logger.info(f"🌍 Region: {args.region}")

        # Optionally get the service principal identity
        if client_id:
            try:
                # Get service principal by appId (client_id)

                query_params = ServicePrincipalsRequestBuilder.ServicePrincipalsRequestBuilderGetQueryParameters(
                    filter=f"appId eq '{client_id}'", select=["displayName", "appId"]
                )
                request_config = ServicePrincipalsRequestBuilder.ServicePrincipalsRequestBuilderGetRequestConfiguration(
                    query_parameters=query_params
                )
                service_principals = await graph_client.service_principals.get(
                    request_configuration=request_config
                )

                if service_principals and service_principals.value:
                    sp = service_principals.value[0]
                    logger.info(f"🤖 Application: {sp.display_name}")
            except Exception:
                # Silently fail - not critical
                pass
    else:
        # For delegated auth, verify by getting user info
        try:
            user = await graph_client.me.get()
            logger.info(f"🔗 Connected to Microsoft Graph as: {user.display_name}")
            # Cache user object for reuse by modules
            cached_user = user
        except ODataError as e:
            # Inspect error code
            code = getattr(e.error, "code", "Unknown")
            message = getattr(e.error, "message", "No message")

            logger.error(
                f"❌ Failed to connect to Microsoft Graph API. Code: {code} | Message: {message}"
            )

            if code == "InvalidAuthenticationToken":
                logger.error(
                    "🔒 Token is invalid or expired. Please re-authenticate with: `msgraphx auth`."
                )
            return 1
        except Exception as exc:
            logger.error(f"❌ Unexpected error while connecting to Graph API: {exc}")
            return 1

    # Create context object with runtime state
    context = GraphContext(
        graph_client=graph_client,
        is_app_only=is_app_only,
        region=args.region,
        cached_user=cached_user,
    )

    # Handle subcommands
    if args.command in ("sharepoint", "sp"):
        if not (
            hasattr(args, "sp_module")
            and hasattr(args, "sp_command")
            and args.sp_command
        ):
            logger.error(
                "Please specify a SharePoint subcommand (e.g., 'msgraphx sp search')"
            )
            return 1
        return await args.sp_module.run_with_arguments(context, args)

    if args.command in ("aad", "ad"):
        if not (
            hasattr(args, "aad_module")
            and hasattr(args, "aad_command")
            and args.aad_command
        ):
            logger.error(
                "Please specify an Azure AD subcommand (e.g., 'msgraphx aad search admin')"
            )
            return 1
        return await args.aad_module.run_with_arguments(context, args)

    if args.command == "me":
        return await me.run_with_arguments(context, args)

    # If no subcommand provided but we authenticated successfully, just show info and exit
    if not args.command:
        logger.info(
            "✅ Authentication successful. Use a subcommand to perform actions."
        )
        return 0

    return 0
