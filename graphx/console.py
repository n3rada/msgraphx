#!/usr/bin/env python3

# Built-in imports
import argparse
import json
from pathlib import Path

# External library imports
from loguru import logger
from msgraph import GraphServiceClient
from msgraph.generated.models.o_data_errors.o_data_error import ODataError


# Local library imports
from graphx.core import terminal
from graphx.core import auth
from graphx.core import logbook
from graphx.core.tokens import TokenManager
from graphx.modules import storage


COMMANDS = {
    "auth": auth,
    "sp": storage,
}


@logger.catch
async def run() -> int:
    parser = argparse.ArgumentParser(
        prog="graphx",
        add_help=True,
        description="Microsoft Graph eXploitation toolkit.",
    )

    parser.add_argument(
        "--log-level",
        choices=["DEBUG", "INFO", "WARNING", "ERROR", "CRITICAL"],
        default="INFO",
        help="Set the logging level (default: INFO).",
    )

    parser.add_argument(
        "--access-token",
        type=str,
        default=None,
        help="Microsoft Graph JWT token. If not provided, the tool will attempt to read it from .roadtools_auth file.",
    )

    parser.add_argument(
        "--refresh-token",
        type=str,
        default=None,
        help="Microsoft Graph refresh token. If not provided, the tool will attempt to read it from .roadtools_auth file.",
    )

    subparsers = parser.add_subparsers(
        dest="command",
        help="Subcommand to run. Leave empty to open interactive shell.",
    )

    parser.add_argument(
        "--before",
        type=str,
        help="Filter items created on or before this date/time. Format: YYYY-MM-DD or relative (e.g. 5h, 3d, 1w).",
    )

    parser.add_argument(
        "--after",
        type=str,
        help="Filter items created on or after this date/time. Format: YYYY-MM-DD or relative (e.g. 5h, 3d, 1w).",
    )

    for name, module in COMMANDS.items():
        subparser = subparsers.add_parser(
            name,
            help=module.__doc__.strip() if module.__doc__ else f"{name} command",
        )
        module.add_arguments(subparser)

    args = parser.parse_args()
    logbook.setup_logging(log_level=args.log_level)

    if args.command == "auth":
        return await auth.run_with_arguments(args)

    access_token = args.access_token
    refresh_token = args.refresh_token

    if not access_token:
        logger.info("🔐 No JWT provided, checking for .roadtools_auth file.")
        tokens_path = Path(".roadtools_auth")
        if tokens_path.exists():
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

        else:
            logger.error("🔐 No JWT provided and .roadtools_auth file not found.")
            return 1

    token = TokenManager(access_token, refresh_token)

    if token.is_expired:
        logger.error("🔒 Token is expired. Please re-authenticate with: `graphx auth`.")
        return 1

    if not (
        token.audience == "00000003-0000-0000-c000-000000000000"
        or token.audience.startswith("https://graph.microsoft.com")
    ):
        logger.error(f"❌ JWT audience mismatch, got '{token.audience}'")
        return 1

    token.start_auto_refresh()

    # Build Graph client
    graph_client = GraphServiceClient(credentials=token)

    try:
        user = await graph_client.me.get()
        logger.info(f"🔗 Connected to Microsoft Graph as: {user.display_name}")
    except ODataError as e:
        # Inspect error code
        code = getattr(e.error, "code", "Unknown")
        message = getattr(e.error, "message", "No message")

        logger.error(
            f"❌ Failed to connect to Microsoft Graph API. Code: {code} | Message: {message}"
        )

        if code == "InvalidAuthenticationToken":
            logger.error(
                "🔒 Token is invalid or expired. Please re-authenticate with: `graphx auth`."
            )
        return 1
    except Exception as e:
        logger.error(f"❌ Unexpected error while connecting to Graph API: {e}")
        return 1
    logger.info(f"🔗 Connected to Microsoft Graph as: {user.display_name}")

    if not args.command:
        logger.info("🔄 Starting interactive shell. Type 'help' for commands.")

        return await terminal.start(graph_client)

    module = COMMANDS[args.command]

    try:
        return await module.run_with_arguments(graph_client=graph_client, args=args)
    except Exception as exc:
        logger.error(f"❌ Error running command '{args.command}': {exc}")
        return 1
