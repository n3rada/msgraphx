#!/usr/bin/env python3

# Built-in imports
import sys
import argparse
import json
from pathlib import Path

# External library imports
from loguru import logger
from modwrap.utils import list_modules

# Local library imports
from graphx.core import logbook
from graphx.core.tokens import TokenManager


MODULES = {
    mod.name: mod
    for mod in list_modules(Path(__file__).parent / "modules")
    if mod.has_callable("run_with_arguments")
}


@logger.catch
def run() -> int:
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
        required=True,
        help="Subcommand to run.",
    )

    for name, wrapper in MODULES.items():
        mod_parser = subparsers.add_parser(name, help=f"{name} module")

        if wrapper.has_callable("add_arguments"):
            wrapper.get_callable("add_arguments")(mod_parser)

        mod_parser.set_defaults(func=wrapper.get_callable("run_with_arguments"))

    if len(sys.argv) == 1:
        parser.print_help()
        return 1

    args = parser.parse_args()
    logbook.setup_logging(log_level=args.log_level)

    if hasattr(args, "func"):
        return args.func(args)

    access_token = args.access_token
    refresh_token = args.refresh_token

    if not access_token:
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

    if not (
        token.audience == "00000003-0000-0000-c000-000000000000"
        or token.audience.startswith("https://graph.microsoft.com")
    ):
        logger.error(f"❌ JWT audience mismatch, got '{token.audience}'")
        return 1

    return 0
