# msgraphx/utils/dates.py

# Built-in imports
import re
from datetime import datetime, timezone, timedelta


def parse_date_string(s: str) -> str:
    """
    Parse string like '5h', '10d', '1w', '2y' or absolute date ('YYYY-MM-DD').
    Returns an ISO 8601 UTC timestamp string (e.g., '2024-12-01T00:00:00Z').
    """
    now = datetime.now(timezone.utc)

    # Relative delta handling
    match = re.match(r"^(\d+)([smhdwy])$", s.lower())
    if match:
        value, unit = match.groups()
        value = int(value)

        delta = {
            "s": timedelta(seconds=value),
            "m": timedelta(minutes=value),
            "h": timedelta(hours=value),
            "d": timedelta(days=value),
            "w": timedelta(weeks=value),
            "y": timedelta(days=value * 365),  # Approximate year as 365 days
        }[unit]

        result_dt = now - delta
        # Format as ISO 8601 with Z suffix (UTC), without microseconds
        return result_dt.strftime("%Y-%m-%dT%H:%M:%SZ")

    # Absolute date handling
    try:
        dt = datetime.strptime(s, "%Y-%m-%d")
        # Assume UTC for date-only inputs
        return dt.strftime("%Y-%m-%dT%H:%M:%SZ")
    except ValueError as exc:
        raise ValueError(
            f"Invalid format: '{s}'. Expected YYYY-MM-DD or duration like 5h, 2d, 1w, 2y."
        ) from exc
