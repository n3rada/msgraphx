# msgraphx/modules/outlook/contacts.py
#
# Build a communication graph from your sent mail: who do you email the most?
# Required delegated permission: Mail.ReadBasic (or Mail.Read)

# Built-in imports
import json
from collections import Counter
from pathlib import Path
from typing import TYPE_CHECKING

# External library imports
from loguru import logger
from msgraph.generated.me.mail_folders.item.messages.messages_request_builder import (
    MessagesRequestBuilder,
)

# Local library imports
from ...utils.errors import handle_graph_errors
from ...utils.dates import parse_date_string

if TYPE_CHECKING:
    import argparse
    from msgraphx.core.context import GraphContext


def add_arguments(parser: "argparse.ArgumentParser"):
    parser.add_argument(
        "--top",
        "-n",
        type=int,
        default=20,
        metavar="N",
        help="Number of top contacts to display (default: 20). Use 0 for all.",
    )

    parser.add_argument(
        "--include-cc",
        action="store_true",
        help="Include CC recipients in the count (default: To only).",
    )


@handle_graph_errors
async def run_with_arguments(
    context: "GraphContext", args: "argparse.Namespace"
) -> int:
    if context.is_app_only:
        logger.error(
            "❌ This module requires delegated authentication (user context)."
        )
        return 1

    # Build OData $filter for date range
    odata_filters: list[str] = []

    if args.after:
        try:
            iso = parse_date_string(args.after)
            odata_filters.append(f"sentDateTime ge {iso}")
        except ValueError as e:
            logger.error(str(e))
            return 1

    if args.before:
        try:
            iso = parse_date_string(args.before)
            odata_filters.append(f"sentDateTime le {iso}")
        except ValueError as e:
            logger.error(str(e))
            return 1

    select_fields = ["toRecipients", "sentDateTime"]
    if args.include_cc:
        select_fields.append("ccRecipients")

    query_params = MessagesRequestBuilder.MessagesRequestBuilderGetQueryParameters(
        select=select_fields,
        top=1000,
    )

    if odata_filters:
        query_params.filter = " and ".join(odata_filters)
        logger.info(f"📅 Date filter: {query_params.filter}")

    request_config = (
        MessagesRequestBuilder.MessagesRequestBuilderGetRequestConfiguration(
            query_parameters=query_params,
        )
    )

    logger.info("📨 Fetching sent messages...")

    to_counter: Counter[str] = Counter()
    cc_counter: Counter[str] = Counter()
    names: dict[str, str] = {}
    total = 0

    result = await context.graph_client.me.mail_folders.by_mail_folder_id(
        "SentItems"
    ).messages.get(request_configuration=request_config)

    while result:
        for msg in result.value or []:
            total += 1

            for r in msg.to_recipients or []:
                if r.email_address and r.email_address.address:
                    addr = r.email_address.address.lower()
                    to_counter[addr] += 1
                    if r.email_address.name:
                        names[addr] = r.email_address.name

            if args.include_cc:
                for r in msg.cc_recipients or []:
                    if r.email_address and r.email_address.address:
                        addr = r.email_address.address.lower()
                        cc_counter[addr] += 1
                        if r.email_address.name:
                            names[addr] = r.email_address.name

        if result.odata_next_link:
            result = await context.graph_client.me.mail_folders.by_mail_folder_id(
                "SentItems"
            ).messages.with_url(result.odata_next_link).get()
        else:
            break

    logger.info(f"📊 Analysed {total:,} sent messages")

    if not to_counter:
        logger.warning("📭 No sent messages found.")
        return 0

    combined: Counter[str] = to_counter.copy()
    if args.include_cc:
        combined.update(cc_counter)

    top_n = args.top if args.top > 0 else None
    ranking = combined.most_common(top_n)

    logger.info(f"🏆 Top {len(ranking)} contacts by sent emails:")

    rank_w = len(str(len(ranking)))
    for rank, (addr, count) in enumerate(ranking, 1):
        name = names.get(addr, "")
        label = f"{name} <{addr}>" if name else addr
        suffix = "s" if count > 1 else ""
        logger.info(
            f"  {rank:>{rank_w}}. {label:<55s}  {count:>5} email{suffix}"
        )

    if args.save:
        save_path = Path(args.save)
        save_path.parent.mkdir(parents=True, exist_ok=True)

        data = [
            {
                "rank": rank,
                "email": addr,
                "name": names.get(addr),
                "to_count": to_counter.get(addr, 0),
                "cc_count": cc_counter.get(addr, 0) if args.include_cc else 0,
                "total_count": count,
            }
            for rank, (addr, count) in enumerate(combined.most_common(), 1)
        ]

        with open(save_path, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2, ensure_ascii=False)

        logger.info(f"💾 Full results saved to: {save_path.absolute()}")

    return 0
