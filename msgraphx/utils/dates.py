import re
from datetime import datetime, timezone, timedelta


def parse_date_string(s: str) -> str:
    """
    Parse string like '5h', '10d', '1w' or absolute date ('YYYY-MM-DD').
    Returns an ISO 8601 UTC timestamp string (e.g., '2024-12-01T00:00:00Z').
    """
    now = datetime.now(timezone.utc)

    # Relative delta handling
    match = re.match(r"^(\d+)([smhdw])$", s.lower())
    if match:
        value, unit = match.groups()
        value = int(value)

        delta = {
            "s": timedelta(seconds=value),
            "m": timedelta(minutes=value),
            "h": timedelta(hours=value),
            "d": timedelta(days=value),
            "w": timedelta(weeks=value),
        }[unit]

        return (now - delta).isoformat() + "Z"

    # Absolute date handling
    try:
        dt = datetime.strptime(s, "%Y-%m-%d")
        return dt.isoformat() + "Z"
    except ValueError as exc:
        raise ValueError(
            f"Invalid format: '{s}'. Expected YYYY-MM-DD or duration like 5h, 2d, 1w."
        ) from exc
