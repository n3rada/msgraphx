# msgraphx/utils/console.py

"""Shared Rich console configured to always use full terminal width."""

import shutil

from rich.console import Console

# Force full terminal width. Rich normally auto-detects, but in some
# environments (piped output, IDE terminals) it may fall back to 80 cols.
# Using shutil.get_terminal_size() as the explicit width ensures we always
# fill the available space.
console = Console(width=shutil.get_terminal_size().columns)
