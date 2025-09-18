#!/usr/bin/env python3

# Standard library imports
import sys

# External library imports
from loguru import logger

# Local library imports
from msgraphx import console

if __name__ == "__main__":
    try:
        sys.exit(console.run())
    except KeyboardInterrupt:
        logger.debug("🛑 User interrupted the process.")
        sys.exit(0)
    except Exception:
        logger.exception("❌ Unexpected exception:")
        sys.exit(1)
