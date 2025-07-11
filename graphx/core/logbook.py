# Built-in imports
import sys
import os

# External library imports
from loguru import logger


# Define the log format
LOG_FORMAT = (
    "<green>{time:YYYY-MM-DD HH:mm:ss.SSS}</green> | "
    "<level>{level: <8}</level> | "
    "<level>{message}</level>"
)


def setup_logging(log_level: str = "INFO"):
    log_level = log_level.upper()

    if not logger.level(log_level, None):
        logger.error(f"Invalid log level: {log_level}")
        logger.warning("Using default log level: INFO")
        log_level = "INFO"

    os.environ["LOG_LEVEL"] = log_level

    # Remove all Loguru handlers to avoid duplicates
    logger.remove()

    # Add a new Loguru handler with custom formatting
    logger.add(
        sys.stderr,
        enqueue=True,
        backtrace=True,
        level=log_level,
        format=LOG_FORMAT,
    )

    logger.success(f"Logging initialized with level {log_level}")
