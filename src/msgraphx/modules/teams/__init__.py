# Built-in imports
from __future__ import annotations

import argparse

# Local library imports
from . import channel, chat, contacts, meetings, send, show
from ...core.context import GraphContext
from ...utils.errors import handle_graph_errors


def add_arguments(
    parser: "argparse.ArgumentParser", parents: "list | None" = None
) -> None:
    parents = parents or []
    subparsers = parser.add_subparsers(dest="subcommand", required=True)

    channel_parser = subparsers.add_parser(
        "channel",
        aliases=["channels"],
        parents=parents,
        help="Search Teams channel messages.",
    )
    channel.add_arguments(channel_parser)

    chat_parser = subparsers.add_parser(
        "chat",
        aliases=["chats"],
        parents=parents,
        help="Search personal chat messages.",
    )
    chat.add_arguments(chat_parser)

    contacts_parser = subparsers.add_parser(
        "contacts", parents=parents, help="Search Teams contacts."
    )
    contacts.add_arguments(contacts_parser)

    send_parser = subparsers.add_parser(
        "send", parents=parents, help="Send a Teams message."
    )
    send.add_arguments(send_parser)

    show_parser = subparsers.add_parser(
        "show", parents=parents, help="Show or browse cached Teams messages."
    )
    show.add_arguments(show_parser)

    meetings_parser = subparsers.add_parser(
        "meetings", parents=parents, help="List online meetings and fetch transcripts."
    )
    meetings.add_arguments(meetings_parser)


@handle_graph_errors
async def run_with_arguments(
    context: "GraphContext", args: "argparse.Namespace"
) -> int:
    subcommand = args.subcommand

    if subcommand in ("channel", "channels"):
        return await channel.run_with_arguments(context, args)
    if subcommand in ("chat", "chats"):
        return await chat.run_with_arguments(context, args)
    if subcommand == "contacts":
        return await contacts.run_with_arguments(context, args)
    if subcommand == "send":
        return await send.run_with_arguments(context, args)
    if subcommand == "show":
        return await show.run_with_arguments(context, args)
    if subcommand == "meetings":
        return await meetings.run_with_arguments(context, args)

    return 1
