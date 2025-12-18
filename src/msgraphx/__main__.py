#!/usr/bin/env python3

# Standard library imports
import asyncio
import sys

# External library imports
from loguru import logger

# Local library imports
from . import cli

if __name__ == "__main__":
    try:
        sys.exit(asyncio.run(cli.main()))
    except KeyboardInterrupt:
        logger.debug("🛑 User interrupted the process.")
        sys.exit(130)
    except Exception:
        logger.exception("❌ Unexpected exception:")
        sys.exit(1)
