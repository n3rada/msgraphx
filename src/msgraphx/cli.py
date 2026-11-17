# msgraphx/cli.py

# Built-in imports
from __future__ import annotations

import argparse
import asyncio
import json
import os
import shutil
import sys
import time
from collections.abc import Awaitable, Callable, Sequence
from pathlib import Path

# External library imports
import httpx
from azure.identity.aio import ClientSecretCredential
from kiota_abstractions.base_request_configuration import RequestConfiguration
from loguru import logger
from rich.console import Console

from msgraph.graph_service_client import GraphServiceClient
from msgraph.generated.service_principals.service_principals_request_builder import (
    ServicePrincipalsRequestBuilder,
)
from msgraph.generated.models.o_data_errors.o_data_error import ODataError
from msgraph.generated.models.user import User

# Local library imports
from . import __version__
from .core.context import GraphContext
from .modules import aad, graph, groups, me, outlook, sharepoint, teams
from .modules.graph import refresh
from .utils import logbook, tokens
from .utils.errors import AuthenticationError, ForbiddenGraphError


def build_parser() -> argparse.ArgumentParser:
    """Build the msgraphx argument parser with every subcommand.

    Global options live on a parent parser shared by all subcommands, so
    they are accepted before or after the subcommand name.

    Returns:
        The top-level parser.
    """
    # Create parent parser with global options that all subcommands inherit.
    # argument_default=SUPPRESS ensures subparsers don't reset values to defaults
    # when a flag is placed before the subcommand (e.g. `msgraphx --debug teams chats`).
    parent_parser = argparse.ArgumentParser(
        add_help=False, argument_default=argparse.SUPPRESS
    )

    advanced_group = parent_parser.add_argument_group(
        "Advanced options", "Logging and debugging controls."
    )

    advanced_group.add_argument(
        "--debug",
        action="store_true",
        default=argparse.SUPPRESS,
        help="Enable debug logging (shortcut for --log-level DEBUG).",
    )

    advanced_group.add_argument(
        "--trace",
        action="store_true",
        default=argparse.SUPPRESS,
        help="Enable TRACE logging (shortcut for --log-level TRACE).",
    )

    advanced_group.add_argument(
        "--log-level",
        type=str,
        choices=["TRACE", "DEBUG", "INFO", "WARNING", "ERROR", "CRITICAL"],
        default=argparse.SUPPRESS,
        help="Set the logging level explicitly.",
    )

    filter_group = parent_parser.add_argument_group(
        "Filters", "Narrow results by time range."
    )

    filter_group.add_argument(
        "--before",
        type=str,
        default=argparse.SUPPRESS,
        help="Filter items created on or before this date/time. Format: YYYY-MM-DD or relative (e.g. 5h, 3d, 1w, 2y).",
    )

    filter_group.add_argument(
        "--after",
        type=str,
        default=argparse.SUPPRESS,
        help="Filter items created on or after this date/time. Format: YYYY-MM-DD or relative (e.g. 5h, 3d, 1w, 2y). Defaults to 1y.",
    )

    filter_group.add_argument(
        "--all",
        action="store_true",
        default=argparse.SUPPRESS,
        dest="fetch_all",
        help="Fetch all items with no time bound (overrides the default --after 1y).",
    )

    output_group = parent_parser.add_argument_group(
        "Output", "Control where results are saved."
    )

    output_group.add_argument(
        "--save",
        "--output",
        "-o",
        type=str,
        default=argparse.SUPPRESS,
        metavar="PATH",
        help="Directory path to save downloaded files. Creates the directory if it doesn't exist.",
    )

    output_group.add_argument(
        "--json",
        action="store_true",
        default=argparse.SUPPRESS,
        dest="json_output",
        help="Output results as JSON to stdout. Suppresses console rendering; logs still go to stderr.",
    )

    output_group.add_argument(
        "--ndjson",
        action="store_true",
        default=argparse.SUPPRESS,
        dest="ndjson_output",
        help="Stream results as NDJSON (one JSON object per line). Ideal for piping to jq or LLM tools.",
    )

    parser = argparse.ArgumentParser(
        prog="msgraphx",
        add_help=True,
        description="Microsoft Graph eXploitation toolkit.",
        parents=[parent_parser],
    )

    # Set proper defaults on the main parser. Subparsers use SUPPRESS so they
    # never overwrite these values when the flag is not present in their argv slice.
    parser.set_defaults(
        debug=False,
        trace=False,
        log_level=None,
        before=None,
        after=None,
        fetch_all=False,
        save=None,
        json_output=False,
        ndjson_output=False,
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
        help="Entra ID Tenant ID for client credentials auth. If not provided, checks TENANT_ID env var.",
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
        help="Entra ID Application (Client) ID (used with --tenant-id). If not provided, checks CLIENT_ID env var.",
    )

    auth_group.add_argument(
        "--client-secret",
        type=str,
        default=None,
        help="Entra ID Client Secret (used with --tenant-id). If not provided, checks CLIENT_SECRET env var.",
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

    # SharePoint subcommand
    sp_parser = subparsers.add_parser(
        "sharepoint",
        aliases=["sp"],
        help="SharePoint commands",
        parents=[parent_parser],
    )
    sharepoint.add_arguments(sp_parser, parents=[parent_parser])

    # Groups subcommand
    groups_parser = subparsers.add_parser(
        "groups",
        help="M365 Unified groups: list, members, sites.",
        parents=[parent_parser],
    )
    groups.add_arguments(groups_parser, parents=[parent_parser])

    # Entra ID subcommand
    aad_parser = subparsers.add_parser(
        "aad",
        aliases=["ad"],
        help="Entra ID commands",
        parents=[parent_parser],
    )
    aad.add_arguments(aad_parser, parents=[parent_parser])

    # Me subcommand
    me_parser = subparsers.add_parser(
        "me",
        help="Current user information",
        parents=[parent_parser],
    )
    me.add_arguments(me_parser)

    # Outlook subcommand
    outlook_parser = subparsers.add_parser(
        "outlook",
        aliases=["mail"],
        help="Outlook / mail commands",
        parents=[parent_parser],
    )
    outlook.add_arguments(outlook_parser, parents=[parent_parser])

    # Teams subcommand
    teams_parser = subparsers.add_parser(
        "teams",
        aliases=["ms-teams"],
        help="Microsoft Teams commands",
        parents=[parent_parser],
    )
    teams.add_arguments(teams_parser, parents=[parent_parser])

    # Token refresh daemon subcommand (runs before auth, no GraphContext needed)
    refresh_parser = subparsers.add_parser(
        "refresh",
        help="Background token refresh daemon to keep delegated sessions alive.",
    )
    refresh.add_arguments(refresh_parser)

    return parser


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _configure_logging(args: argparse.Namespace) -> None:
    """Set up logging at the level chosen on the command line.

    `--log-level` wins over `--debug`, which wins over `--trace`; the
    default is INFO.

    Args:
        args: Parsed arguments with `log_level`, `debug`, and `trace`.
    """
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
    """Route all HTTP traffic through a proxy via the proxy env variables.

    Args:
        proxy: Proxy URL starting with `http://` or `https://`, or `None`
            for no proxy.

    Returns:
        1 if the proxy URL has an unsupported scheme, otherwise 0.
    """
    if not proxy:
        return 0
    if not proxy.startswith(("http://", "https://")):
        logger.error("Invalid proxy format.")
        return 1
    for key in (
        "http_proxy",
        "https_proxy",
        "HTTP_PROXY",
        "HTTPS_PROXY",
        "ALL_PROXY",
        "all_proxy",
    ):
        os.environ[key] = proxy
    logger.info(f"Proxy set to {proxy}")
    return 0


def _check_public_ip() -> tuple[str | None, float]:
    """Look up the public egress IP via api.ipify.org, through any proxy.

    Returns:
        A tuple (ip, rtt), where ip is the public IP address or `None` if the
        lookup failed and rtt is the request round-trip time in seconds
        (0.0 on failure).
    """
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
        logger.warning(f"Error retrieving public IP: {exc}")
    return None, 0.0


def _load_token(args: argparse.Namespace) -> tuple[str | None, str | None, str]:
    """Find the delegated tokens on the command line, env, or token file.

    Checked in order: `--access-token`/`--refresh-token`, the
    `ACCESS_TOKEN`/`REFRESH_TOKEN` environment variables, then
    `.roadtools_auth` in the working directory.

    Args:
        args: Parsed arguments with `access_token` and `refresh_token`.

    Returns:
        A tuple (access_token, refresh_token, source), where source is
        `arg`, `env`, or `file` and decides where refreshed tokens are
        written back. access_token is `None` if no token was found or the
        token file is not valid JSON.
    """
    if args.access_token:
        return args.access_token, args.refresh_token, "arg"

    env_token = os.environ.get("ACCESS_TOKEN")
    if env_token:
        logger.info("Using JWT from environment variable ACCESS_TOKEN.")
        return env_token, os.environ.get("REFRESH_TOKEN"), "env"

    logger.info("No JWT provided, checking for .roadtools_auth file.")
    tokens_path = Path(".roadtools_auth")
    if not tokens_path.exists():
        logger.error("No JWT provided and .roadtools_auth file not found.")
        return None, None, "file"

    logger.info("Using JWT from .roadtools_auth file.")
    try:
        data = json.loads(tokens_path.read_bytes())
        return data.get("accessToken"), data.get("refreshToken"), "file"
    except json.JSONDecodeError:
        logger.error(
            "Failed to decode .roadtools_auth file. Ensure it contains valid JSON."
        )
        return None, None, "file"


def _build_app_client(
    tenant_id: str, client_id: str, client_secret: str
) -> tuple[GraphServiceClient | None, ClientSecretCredential | None]:
    """Create a Graph client that authenticates with a client secret.

    No token is requested yet; credential errors surface on the first call.

    Args:
        tenant_id: Directory (tenant) ID.
        client_id: Application (client) ID.
        client_secret: Client secret value.

    Returns:
        A tuple (client, credentials), or `(None, None)` if construction
        failed, which is logged.
    """
    try:
        credentials = ClientSecretCredential(
            tenant_id=tenant_id,
            client_id=client_id,
            client_secret=client_secret,
        )
        client = GraphServiceClient(
            credentials=credentials,
            scopes=["https://graph.microsoft.com/.default"],
        )
        return client, credentials
    except Exception as exc:
        logger.exception("Failed to create Graph client with client credentials")
        return None, None


def _classify_token(token: tokens.TokenManager) -> bool:
    """Check whether a token carries app-only permissions, and log which.

    Args:
        token: The decoded access token.

    Returns:
        True if the token has a `roles` claim and no `scp` claim.
    """
    payload = token.payload
    if payload.get("roles") and not payload.get("scp"):
        logger.info("Token contains application permissions (app roles).")
        return True
    logger.info("Token contains delegated permissions (user scopes).")
    return False


async def _authenticate(
    args: argparse.Namespace,
) -> tuple[
    GraphServiceClient | None,
    bool,
    frozenset[str],
    Callable[[], Awaitable[str]] | None,
    int,
]:
    """Choose the authentication method and build the Graph client.

    Client credentials are used with `--tenant-id`, or when
    `TENANT_ID`/`CLIENT_ID`/`CLIENT_SECRET` are all set and no access token
    was given. Otherwise a delegated token is loaded with `_load_token`,
    checked for expiry and Graph audience, and auto-refreshed in a
    background thread unless the refresh daemon is running. Resolved client
    credentials are written back onto `args`.

    Args:
        args: Parsed arguments with the authentication options.

    Returns:
        A tuple (client, is_app_only, token_scopes, token_getter,
        exit_code). On failure, client and token_getter are `None` and
        exit_code is 1; on success exit_code is 0.

    Raises:
        ValueError: If the delegated access token is not a decodable JWT.
    """
    if args.access_token and args.tenant_id:
        logger.error(
            "Cannot use both token-based authentication (--access-token) and "
            "client credentials (--tenant-id) simultaneously."
        )
        return None, False, frozenset(), None, 1

    has_env_creds = all(
        os.environ.get(k) for k in ("TENANT_ID", "CLIENT_ID", "CLIENT_SECRET")
    )
    use_creds = args.tenant_id is not None or (
        has_env_creds
        and args.access_token is None
        and not os.environ.get("ACCESS_TOKEN")
    )

    if use_creds:
        tenant_id = args.tenant_id or os.environ.get("TENANT_ID")
        client_id = args.client_id or os.environ.get("CLIENT_ID")
        client_secret = args.client_secret or os.environ.get("CLIENT_SECRET")
        args.tenant_id, args.client_id, args.client_secret = (
            tenant_id,
            client_id,
            client_secret,
        )

        if tenant_id:
            logger.info(f"Tenant ID: {tenant_id}")
        if client_id:
            logger.info(f"Client ID: {client_id}")

        logger.info("Using client credentials authentication")
        client, credentials = _build_app_client(tenant_id, client_id, client_secret)
        if not client:
            return None, True, frozenset(), None, 1

        async def _app_token_getter() -> str:
            t = await credentials.get_token("https://graph.microsoft.com/.default")
            return t.token

        return client, True, frozenset(), _app_token_getter, 0

    access_token, refresh_token, token_source = _load_token(args)
    if not access_token:
        return None, False, frozenset(), None, 1

    token = tokens.TokenManager(access_token, refresh_token, source=token_source)

    if token.is_expired:
        logger.error("Token is expired.")
        return None, False, frozenset(), None, 1

    if (
        token.audience != "00000003-0000-0000-c000-000000000000"
        and not token.audience.startswith("https://graph.microsoft.com")
    ):
        logger.error(f"JWT audience mismatch, got '{token.audience}'")
        return None, False, frozenset(), None, 1

    is_app_only = _classify_token(token)

    token_scopes: frozenset[str] = frozenset()
    if token.app_id:
        logger.info(f"Token app: {token.app_id}")
    if not is_app_only and token.scope:
        scope_list = token.scope.split()
        token_scopes = frozenset(scope_list)
        logger.debug(f"Token scopes ({len(scope_list)}): {', '.join(scope_list)}")

    if refresh.is_running():
        logger.debug("Background refresh daemon detected — skipping inline refresh.")
    else:
        token.start_auto_refresh()

    async def _delegated_token_getter() -> str:
        return token.access_token

    return GraphServiceClient(token), is_app_only, token_scopes, _delegated_token_getter, 0


async def _verify_connection(
    graph_client: GraphServiceClient,
    is_app_only: bool,
    client_id: str | None,
    region: str,
) -> tuple[User | None, int]:
    """Check the Graph connection and log who is authenticated.

    Delegated tokens are verified with `GET /me`. App-only clients cannot
    call `/me`, so they only log the region and, when possible, the
    application's display name.

    Args:
        graph_client: The authenticated Graph client.
        is_app_only: Whether the client uses app-only permissions.
        client_id: Application ID used to look up the app's name, or `None`.
        region: Search region, logged for app-only clients.

    Returns:
        A tuple (user, exit_code), where user is the signed-in user (`None`
        for app-only or on failure) and exit_code is 0 on success or 1 if
        the connection failed, which is logged.
    """
    if is_app_only:
        logger.info("Authenticated with application permissions.")
        logger.info(f"Region: {region}")
        if client_id:
            await _log_service_principal(graph_client, client_id)
        return None, 0

    try:
        user = await graph_client.me.get()
        logger.info(
            f"Authenticated as: {user.display_name} ({user.user_principal_name})"
        )
        return user, 0
    except ODataError as exc:
        code = getattr(exc.error, "code", "Unknown")
        message = getattr(exc.error, "message", "No message")
        logger.error(
            f"Failed to connect to Microsoft Graph API. Code: {code} | Message: {message}"
        )
        if code == "InvalidAuthenticationToken":
            logger.error("Token is invalid or expired. Please re-authenticate.")
        return None, 1
    except TimeoutError:
        logger.error("Connection timeout. Check your network connection.")
        return None, 1
    except Exception as exc:
        logger.exception("Unexpected error while connecting to Graph API")
        return None, 1


async def _log_service_principal(
    graph_client: GraphServiceClient, client_id: str
) -> None:
    """Log the display name of the service principal for an application.

    Lookup failures are logged, not raised.

    Args:
        graph_client: The authenticated Graph client.
        client_id: The application (client) ID to look up.
    """
    try:
        query_params = ServicePrincipalsRequestBuilder.ServicePrincipalsRequestBuilderGetQueryParameters(
            filter=f"appId eq '{client_id}'", select=["displayName", "appId"]
        )
        request_config = RequestConfiguration(query_parameters=query_params)
        sps = await graph_client.service_principals.get(
            request_configuration=request_config
        )
        if sps and sps.value:
            logger.info(f"Application: {sps.value[0].display_name}")
    except Exception as exc:
        logger.error("Failed to log service principal details")


async def _call_module(coro: Awaitable[int]) -> int:
    """Await a module entry point and turn Graph access errors into exit codes.

    A `ForbiddenGraphError` is printed to stderr with the required and
    granted permissions. An `AuthenticationError` is assumed to be logged
    already.

    Args:
        coro: The pending `run_with_arguments` call.

    Returns:
        The module's exit code, or 1 on a 403 or authentication failure.
    """
    try:
        return await coro
    except ForbiddenGraphError as exc:
        console = Console(stderr=True, width=shutil.get_terminal_size().columns)
        console.print("\n[bold red]Access denied (403 Forbidden)[/bold red]")
        if exc.required:
            console.print(f"  [bold]Required:[/bold]  {exc.required}")
        if exc.granted:
            console.print(f"  [bold]Granted:[/bold]   {exc.granted}")
        if not exc.required and exc.raw_message:
            console.print(f"  {exc.raw_message}")
        return 1
    except AuthenticationError:
        # Already logged by the decorator
        return 1


async def _dispatch(args: argparse.Namespace, context: GraphContext) -> int:
    """Run the module for the selected subcommand.

    Args:
        args: Parsed arguments; `command` selects the module.
        context: The shared Graph context.

    Returns:
        The module's exit code, or 0 when no subcommand was given.
    """
    command = args.command

    if command == "groups":
        return await _call_module(groups.run_with_arguments(context, args))

    if command in ("sharepoint", "sp"):
        return await _call_module(sharepoint.run_with_arguments(context, args))

    if command in ("aad", "ad"):
        return await _call_module(aad.run_with_arguments(context, args))

    if command == "me":
        return await _call_module(me.run_with_arguments(context, args))

    if command in ("outlook", "mail"):
        return await _call_module(outlook.run_with_arguments(context, args))

    if command in ("teams", "ms-teams"):
        return await _call_module(teams.run_with_arguments(context, args))

    logger.info("Authentication successful. Use a subcommand to perform actions.")
    return 0


# ---------------------------------------------------------------------------
# Entry points
# ---------------------------------------------------------------------------


def _pre_parse_globals(argv: Sequence[str] | None = None) -> argparse.Namespace:
    """Parse the global flags on their own, ignoring everything else.

    Works around argparse nested subparsers overwriting values set by a
    parent parser, even with `SUPPRESS` defaults.

    Args:
        argv: Command-line arguments without the program name, or `None`
            for `sys.argv[1:]`.

    Returns:
        A namespace with every global flag; unset flags are `None` or
        False.
    """
    p = argparse.ArgumentParser(add_help=False)
    p.add_argument("--debug", action="store_true", default=False)
    p.add_argument("--trace", action="store_true", default=False)
    p.add_argument(
        "--log-level",
        choices=["TRACE", "DEBUG", "INFO", "WARNING", "ERROR", "CRITICAL"],
        default=None,
    )
    p.add_argument("--before", type=str, default=None)
    p.add_argument("--after", type=str, default=None)
    p.add_argument("--all", dest="fetch_all", action="store_true", default=False)
    p.add_argument("--save", "--output", "-o", type=str, default=None)
    p.add_argument("--json", dest="json_output", action="store_true", default=False)
    p.add_argument("--ndjson", dest="ndjson_output", action="store_true", default=False)
    ns, _ = p.parse_known_args(argv)
    return ns


@logger.catch
async def _main(argv: Sequence[str]) -> int:
    """Run the msgraphx command line.

    Parses arguments, configures logging and proxying, checks the public
    IP, authenticates, verifies the connection, and dispatches the
    subcommand. Commands that use time bounds default to `--after 1y`
    unless `--all` is given. Because of `logger.catch`, an unexpected
    exception is logged and the call returns `None` instead of raising.

    Args:
        argv: Command-line arguments without the program name.

    Returns:
        The process exit code.
    """
    # Pre-parse global flags to survive nested subparser default-stomping.
    # argparse resets parent-parser values when nested subparsers are involved;
    # _pre_parse_globals extracts globals from the raw argv independently.
    globals_ns = _pre_parse_globals(argv)

    parser = build_parser()
    args = parser.parse_args(argv)

    # Merge globals back: only override values that were explicitly passed
    # (non-None for strings, True for store_true flags; never False, never None).
    for key, value in vars(globals_ns).items():
        if value is not None and value is not False:
            setattr(args, key, value)

    _configure_logging(args)
    logger.trace(
        f"debug={args.debug!r}  trace={args.trace!r}  log_level={args.log_level!r}"
    )

    if args.command == "refresh":
        return refresh.run_command(args)

    if _apply_proxy(getattr(args, "proxy", None)):
        return 1

    public_ip, rtt = _check_public_ip()
    if not public_ip:
        return 1

    logger.info(f"Public IP: {public_ip} (RTT: {rtt:.2f}s)")

    # Only apply a default --after time bound for subcommands that opted in
    # via parser.set_defaults(uses_time_bounds=True) in their add_arguments().
    if getattr(args, "uses_time_bounds", False):
        if not getattr(args, "fetch_all", False) and not args.after:
            args.after = "1y"
            logger.info(
                "No time bound specified, defaulting to last year. Use --all to fetch everything."
            )

    args.drive_id = getattr(args, "drive_id", None) or os.environ.get("DRIVE_ID")
    if args.drive_id:
        logger.info(f"Drive ID: {args.drive_id}")

    graph_client, is_app_only, token_scopes, token_getter, err = await _authenticate(args)
    if err:
        return err

    args.is_app_only = is_app_only

    cached_user, err = await _verify_connection(
        graph_client, is_app_only, getattr(args, "client_id", None), args.region
    )
    if err:
        return err

    # Build raw identity for cache partitioning.
    # Delegated: OID comes from /me (cached_user.id set by _verify_connection).
    # App-only: /me is unavailable; use tenant_id:client_id as the surrogate.
    if is_app_only:
        raw_identity = f"{getattr(args, 'tenant_id', '') or ''}:{getattr(args, 'client_id', '') or ''}"
    else:
        raw_identity = ""  # GraphContext.identity_hash falls back to cached_user.id

    context = GraphContext(
        graph_client=graph_client,
        is_app_only=is_app_only,
        region=args.region,
        cached_user=cached_user,
        token_scopes=token_scopes,
        json_output=getattr(args, "json_output", False),
        ndjson_output=getattr(args, "ndjson_output", False),
        token_getter=token_getter,
        raw_identity=raw_identity,
    )

    return await _dispatch(args, context)


def main() -> None:
    """Run msgraphx with `sys.argv` and exit with its status.

    Raises:
        SystemExit: Always, with the command's exit code, or 130 when
            interrupted with Ctrl+C.
    """
    try:
        raise SystemExit(asyncio.run(_main(sys.argv[1:])))
    except KeyboardInterrupt:
        raise SystemExit(130)
