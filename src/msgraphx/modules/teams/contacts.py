# msgraphx/modules/teams/contacts.py
#
# Build a communication graph from your Teams chats.
# Lists 1:1 chat partners and group chat participants.
# Required delegated permissions: Chat.Read, Chat.ReadBasic
#
# GET /me/chats?$expand=members gives all chats + their members in one call.
# For message frequency, use GET /me/chats/{id}/messages (expensive — opt-in via --count).
#
# Tip: prototype in Graph Explorer
# https://developer.microsoft.com/en-us/graph/graph-explorer

# Built-in imports
from __future__ import annotations

import argparse
import json
from collections import Counter
from pathlib import Path

# External library imports
from kiota_abstractions.base_request_configuration import RequestConfiguration
from loguru import logger
from msgraph.generated.users.item.chats.chats_request_builder import ChatsRequestBuilder
from rich.console import Console
from rich.table import Table

# Local library imports
from ...core.context import GraphContext
from ...utils.errors import handle_graph_errors
from ...utils.pagination import GraphPaginator


def add_arguments(parser: "argparse.ArgumentParser") -> None:
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
        help="Fetch message counts per chat (slower — makes one extra API call per chat).",
    )


@handle_graph_errors
async def run_with_arguments(
    context: "GraphContext", args: "argparse.Namespace"
) -> int:
    if context.is_app_only:
        logger.error("This module requires delegated authentication (user context).")
        return 1

    me = await context.graph_client.me.get()
    my_id = me.id if me else None

    logger.info("💬 Fetching Teams chats...")

    query_params = ChatsRequestBuilder.ChatsRequestBuilderGetQueryParameters(
        expand=["members"],
        top=50,
    )
    request_config = RequestConfiguration(query_parameters=query_params)

    dm_counter: Counter[str] = Counter()
    group_counter: Counter[str] = Counter()
    names: dict[str, str] = {}

    one_on_one: list[dict] = []
    groups: list[dict] = []

    async for chat in GraphPaginator(context.graph_client.me.chats, request_config):
        chat_type = str(chat.chat_type) if chat.chat_type else ""
        members = chat.members or []

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
                    if (getattr(m, "user_id", None) or getattr(m, "id", None)) != my_id
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

    top_n = args.top if args.top > 0 else None
    console = Console()

    # 1:1 chat partners ranked by chat count (proxy for communication frequency)
    dm_ranking = dm_counter.most_common(top_n)
    if dm_ranking:
        table = Table(
            title=f"💬 Top {args.top or len(dm_ranking)} — Direct message partners",
            show_header=True,
            header_style="bold",
        )
        table.add_column("#", justify="right", style="dim")
        table.add_column("Chats", justify="right")
        table.add_column("Contact")
        for rank, (uid, cnt) in enumerate(dm_ranking, 1):
            table.add_row(str(rank), str(cnt), names.get(uid, uid))
        console.print(table)

    # Group chats
    group_ranking = group_counter.most_common(top_n)
    if group_ranking:
        table = Table(
            title=f"👥 Top {args.top or len(group_ranking)} — Group chat participants",
            show_header=True,
            header_style="bold",
        )
        table.add_column("#", justify="right", style="dim")
        table.add_column("Chats", justify="right")
        table.add_column("Contact")
        for rank, (uid, cnt) in enumerate(group_ranking, 1):
            table.add_row(str(rank), str(cnt), names.get(uid, uid))
        console.print(table)

    logger.info(
        f"📊 {len(one_on_one)} 1:1 chats, {len(groups)} group chats, "
        f"{len(dm_counter) + len(group_counter)} unique contacts"
    )

    if args.save:
        save_path = Path(args.save)
        save_path.parent.mkdir(parents=True, exist_ok=True)
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
        with open(save_path, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
        logger.info(f"💾 Full results saved to: {save_path.absolute()}")

    return 0
