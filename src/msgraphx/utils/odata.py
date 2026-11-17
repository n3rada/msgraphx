# msgraphx/utils/odata.py

# Built-in imports
from __future__ import annotations


def quote(value: str) -> str:
    """Return `value` as a quoted OData string literal for `$filter`.

    Single quotes are escaped by doubling them, per the OData URL
    conventions, so input cannot break out of the literal.

    Args:
        value: Raw string to embed in a `$filter` expression.

    Returns:
        The literal including surrounding quotes, e.g. `'o''brien@corp.com'`.
    """
    return "'" + value.replace("'", "''") + "'"
