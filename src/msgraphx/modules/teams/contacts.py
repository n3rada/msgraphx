# msgraphx/modules/teams/contacts.py
#
# Build a communication graph from your Teams chats.
# Lists 1:1 chat partners and group chat participants.
# Required delegated permissions: Chat.Read, Chat.ReadBasic
#
# GET /me/chats?$expand=members gives all chats + their members in one call.
# For message frequency, use GET /me/chats/{id}/messages (expensive, opt-in via --count).
#
# Tip: prototype in Graph Explorer
# https://developer.microsoft.com/en-us/graph/graph-explorer

# Built-in imports
from __future__ import annotations

import argparse
import contextlib
import json
from collections import Counter
from datetime import datetime, timezone
from pathlib import Path

# External library imports
from kiota_abstractions.base_request_configuration import RequestConfiguration
from loguru import logger
from msgraph.generated.users.item.chats.chats_request_builder import ChatsRequestBuilder
from rich.console import Group
from rich.live import Live
from rich.table import Table

# Local library imports
from ...core.context import GraphContext
from ...utils import output
from ...utils.console import console
from ...utils.dates import parse_date_string
from ...utils.errors import handle_graph_errors
from ...utils.pagination import GraphPaginator
from ...utils.roles import require_scopes


def add_arguments(parser: "argparse.ArgumentParser") -> None:
    parser.set_defaults(uses_time_bounds=True)
    parser.add_argument(
        "--top",
        "-n",
        type=int,
        default=30,
        metavar="N",
        help="Number of top contacts to display (default: 30). Use 0 for all.",
    )

    parser.add_argument(
        "--count",
        action="store_true",
        help="Fetch message counts per chat (slower, makes one extra API call per chat).",
    )


@handle_graph_errors
@require_scopes("Chat.Read")
async def run_with_arguments(
    context: "GraphContext", args: "argparse.Namespace"
) -> int:
    if context.is_app_only:
        logger.error("This module requires delegated authentication (user context).")
        return 1

    if not (context.has_scope("Chat.Read") or context.has_scope("Chat.ReadBasic")):
        logger.error(
            "Missing required scope. Grant at least Chat.ReadBasic or Chat.Read."
        )
        return 1

    me = await context.graph_client.me.get()
    my_id = me.id if me else None

    logger.info("Fetching Teams chats")

    query_params = ChatsRequestBuilder.ChatsRequestBuilderGetQueryParameters(
        expand=["members", "lastMessagePreview"],
        orderby=["lastMessagePreview/createdDateTime desc"],
        top=50,
    )
    request_config = RequestConfiguration(query_parameters=query_params)

    after_cutoff: datetime | None = None
    if args.after:
        try:
            after_cutoff = datetime.fromisoformat(
                parse_date_string(args.after)
            ).replace(tzinfo=timezone.utc)
        except ValueError as e:
            logger.error(str(e))
            return 1

    dm_counter: Counter[str] = Counter()
    group_counter: Counter[str] = Counter()
    names: dict[str, str] = {}

    one_on_one: list[dict] = []
    groups: list[dict] = []
    total = 0

    top_n = args.top if args.top > 0 else None

    def _build_table(counter: Counter, title: str) -> Table:
        table = Table(title=title, show_header=True, header_style="bold")
        table.add_column("#", justify="right", style="dim")
        table.add_column("Chats", justify="right")
        table.add_column("Contact")
        for rank, (uid, cnt) in enumerate(counter.most_common(top_n), 1):
            table.add_row(str(rank), str(cnt), names.get(uid, uid))
        return table

    live_ctx = contextlib.nullcontext() if context.json_output else Live(console=console, refresh_per_second=4)
    with live_ctx as live:
        async for chat in GraphPaginator(context.graph_client.me.chats, request_config):
            total += 1
            if live is not None:
                live.update(
                    Group(
                        _build_table(
                            dm_counter,
                            f"Top {args.top or len(dm_counter)} : Direct message partners ({total:,} scanned)",
                        ),
                        _build_table(
                            group_counter,
                            f"Top {args.top or len(group_counter)} : Group chat participants",
                        ),
                    )
                )

            chat_type = str(chat.chat_type) if chat.chat_type else ""
            members = chat.members or []

            # Chats are ordered newest-first; stop once we pass the lower date bound
            if after_cutoff and chat.last_message_preview:
                last_ts = chat.last_message_preview.created_date_time
                if last_ts:
                    if last_ts.tzinfo is None:
                        last_ts = last_ts.replace(tzinfo=timezone.utc)
                    if last_ts < after_cutoff:
                        break

            for m in members:
                uid = getattr(m, "user_id", None) or getattr(m, "id", None)
                display = getattr(m, "display_name", None) or ""
                if uid and uid != my_id:
                    if display:
                        names[uid] = display
                    if "oneOnOne" in chat_type:
                        dm_counter[uid] += 1
                    else:
                        group_counter[uid] += 1

            if "oneOnOne" in chat_type:
                other = next(
                    (
                        m
                        for m in members
                        if (getattr(m, "user_id", None) or getattr(m, "id", None))
                        != my_id
                    ),
                    None,
                )
                if other:
                    one_on_one.append(
                        {
                            "chat_id": chat.id,
                            "display_name": getattr(other, "display_name", "") or "",
                            "email": getattr(other, "email", "") or "",
                        }
                    )
            else:
                groups.append(
                    {
                        "chat_id": chat.id,
                        "topic": chat.topic or "",
                        "member_count": len(members),
                    }
                )

    logger.info(
        f"{len(one_on_one)} 1:1 chats, {len(groups)} group chats, "
        f"{len(dm_counter) + len(group_counter)} unique contacts"
    )

    if args.save or context.json_output:
        data = {
            "one_on_one": one_on_one,
            "groups": groups,
            "dm_ranking": [
                {"uid": uid, "name": names.get(uid, uid), "chat_count": cnt}
                for uid, cnt in dm_counter.most_common()
            ],
            "group_ranking": [
                {"uid": uid, "name": names.get(uid, uid), "chat_count": cnt}
                for uid, cnt in group_counter.most_common()
            ],
        }

        if args.save:
            save_path = Path(args.save)
            save_path.parent.mkdir(parents=True, exist_ok=True)
            with open(save_path, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=2, ensure_ascii=False)
            logger.info(f"Full results saved to: {save_path.absolute()}")

        if context.json_output:
            output.print_json(data)

    return 0
