# Built-in imports
from pathlib import Path

# External library imports
from loguru import logger
from prompt_toolkit import PromptSession
from prompt_toolkit.auto_suggest import ThreadedAutoSuggest, AutoSuggestFromHistory
from prompt_toolkit.history import ThreadedHistory, InMemoryHistory
from prompt_toolkit.cursor_shapes import CursorShape
from prompt_toolkit.completion import WordCompleter


# Local library imports
from graphx.core import auth, logbook
from graphx.core.tokens import TokenManager
from graphx.core import terminal

commands = ["me", "groups", "search", "exit", "help"]
completer = WordCompleter(commands, ignore_case=True)


async def start(graph_client) -> int:
    session = PromptSession(
        cursor=CursorShape.BLINKING_BLOCK,
        multiline=False,
        enable_history_search=True,
        wrap_lines=True,
        auto_suggest=ThreadedAutoSuggest(auto_suggest=AutoSuggestFromHistory()),
        history=ThreadedHistory(history=InMemoryHistory()),
    )
    while True:
        try:
            cmd_line = await session.prompt_async(
                "graphx> ",
                completer=completer,
            )
        except (EOFError, KeyboardInterrupt):
            break

        parts = cmd_line.strip().split()
        if not parts:
            continue

        cmd, *args = parts

        if cmd == "exit":
            break

        if cmd == "help":
            print("Available:", ", ".join(commands))
        elif cmd == "me":
            user = await graph_client.me.get()
            print(user)
        elif cmd == "groups":
            groups = await graph_client.groups.get()

            print(groups)

        else:
            print(f"❌ Unknown command: {cmd}")

    return 0
