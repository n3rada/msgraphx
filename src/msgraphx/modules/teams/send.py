# msgraphx/modules/teams/send.py
#
# Send a Teams message to a user (1:1 chat) or a channel.
# Required delegated permissions:
#   Chat.ReadWrite: create/use 1:1 chats
#   ChatMessage.Send: send to chats
#   ChannelMessage.Send: send to channels
#
# Usage:
#   teams send alice@corp.com "Hey, lunch today?"
#   teams send --team <teamId> --channel <channelId> "Announcement text"
#
# Tip: prototype in Graph Explorer
# https://developer.microsoft.com/en-us/graph/graph-explorer

# Built-in imports
from __future__ import annotations

import argparse

# External library imports
from loguru import logger
from msgraph.generated.models.aad_user_conversation_member import (
    AadUserConversationMember,
)
from msgraph.generated.models.body_type import BodyType
from msgraph.generated.models.chat import Chat
from msgraph.generated.models.chat_message import ChatMessage
from msgraph.generated.models.chat_type import ChatType
from msgraph.generated.models.item_body import ItemBody

# Local library imports
from ...core.context import GraphContext
from ...utils.errors import handle_graph_errors


def add_arguments(parser: "argparse.ArgumentParser") -> None:
    target_group = parser.add_mutually_exclusive_group(required=True)

    target_group.add_argument(
        "--to",
        metavar="USER",
        type=str,
        help="Recipient UPN or object ID (for 1:1 chat).",
    )

    target_group.add_argument(
        "--channel",
        metavar="CHANNEL_ID",
        type=str,
        help="Channel ID to post to (requires --team).",
    )

    parser.add_argument(
        "--team",
        metavar="TEAM_ID",
        type=str,
        default=None,
        help="Team ID, required when --channel is specified.",
    )

    parser.add_argument(
        "message",
        type=str,
        help="Message body (plain text).",
    )

    parser.add_argument(
        "--html",
        action="store_true",
        help="Treat the message body as HTML.",
    )


@handle_graph_errors
async def run_with_arguments(
    context: "GraphContext", args: "argparse.Namespace"
) -> int:
    if context.is_app_only:
        logger.error("This module requires delegated authentication (user context).")
        return 1

    body_type = BodyType.Html if args.html else BodyType.Text
    message_body = ItemBody(content=args.message, content_type=body_type)

    if args.channel:
        return await _send_channel_message(context, args, message_body)

    return await _send_dm(context, args, message_body)


async def _send_dm(
    context: "GraphContext",
    args: "argparse.Namespace",
    message_body: "ItemBody",
) -> int:
    """Find or create a 1:1 chat with the target user, then post the message."""
    me = await context.graph_client.me.get()
    if not me or not me.id:
        logger.error("Failed to resolve current user identity.")
        return 1

    recipient = args.to
    logger.info(f"Resolving user: {recipient}")

    # Accept either UPN or object ID; the by_user_id call works for both
    try:
        target_user = await context.graph_client.users.by_user_id(recipient).get()
    except (ValueError, RuntimeError, OSError) as exc:
        logger.error(f"Failed to resolve user '{recipient}': {exc}")
        return 1

    if not target_user or not target_user.id:
        logger.error(f"User not found: {recipient}")
        return 1

    logger.info(
        f"Creating or reusing 1:1 chat with "
        f"{target_user.display_name} ({target_user.mail or target_user.id})"
    )

    chat_request = Chat(
        chat_type=ChatType.OneOnOne,
        members=[
            AadUserConversationMember(
                odata_type="#microsoft.graph.aadUserConversationMember",
                roles=["owner"],
                additional_data={
                    "user@odata.bind": (
                        f"https://graph.microsoft.com/v1.0/users('{me.id}')"
                    )
                },
            ),
            AadUserConversationMember(
                odata_type="#microsoft.graph.aadUserConversationMember",
                roles=["owner"],
                additional_data={
                    "user@odata.bind": (
                        f"https://graph.microsoft.com/v1.0/users('{target_user.id}')"
                    )
                },
            ),
        ],
    )

    chat = await context.graph_client.chats.post(chat_request)
    if not chat or not chat.id:
        logger.error("Failed to create or retrieve 1:1 chat.")
        return 1

    chat_id = chat.id
    logger.debug(f"Chat ID: {chat_id}")

    outgoing = ChatMessage(body=message_body)
    result = await context.graph_client.chats.by_chat_id(chat_id).messages.post(
        outgoing
    )

    if result:
        preview = args.message[:80] + ("…" if len(args.message) > 80 else "")
        logger.success(
            f"Message sent to {target_user.display_name or recipient}: {preview!r}"
        )
    else:
        logger.error("Message POST returned no result.")
        return 1

    return 0


async def _send_channel_message(
    context: "GraphContext",
    args: "argparse.Namespace",
    message_body: "ItemBody",
) -> int:
    """Post a message to a Teams channel."""
    if not args.team:
        logger.error("--team <TEAM_ID> is required when using --channel.")
        return 1

    team_id = args.team
    channel_id = args.channel

    logger.info(f"Posting to channel {channel_id} in team {team_id}")

    outgoing = ChatMessage(body=message_body)
    result = await (
        context.graph_client.teams.by_team_id(team_id)
        .channels.by_channel_id(channel_id)
        .messages.post(outgoing)
    )

    if result:
        preview = args.message[:80] + ("…" if len(args.message) > 80 else "")
        logger.success(f"Message posted to channel: {preview!r}")
    else:
        logger.error("Channel message POST returned no result.")
        return 1

    return 0
