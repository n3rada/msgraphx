# msgraphx/modules/outlook/contacts.py
#
# Build a communication graph from your mail: who do you exchange with most?
# Required delegated permission: Mail.ReadBasic (or Mail.Read)

# Built-in imports
from __future__ import annotations

import json
from collections import Counter
from pathlib import Path

# External library imports
from kiota_abstractions.base_request_configuration import RequestConfiguration
from loguru import logger
from msgraph.generated.users.item.mail_folders.item.messages.messages_request_builder import (
    MessagesRequestBuilder,
)
from rich.console import Console
from rich.table import Table

# Local library imports
from ...utils.errors import handle_graph_errors
from ...utils.dates import parse_date_string

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
        "--all",
        action="store_true",
        dest="fetch_all",
        help="Fetch all sent messages with no time bound (overrides --after default of 1y).",
    )

    parser.add_argument(
        "--only",
        choices=["sent", "received"],
        default=None,
        metavar="{sent,received}",
        help="Restrict analysis to sent or received mail only (default: both).",
    )


@handle_graph_errors
async def run_with_arguments(
    context: "GraphContext", args: "argparse.Namespace"
) -> int:
    if context.is_app_only:
        logger.error("❌ This module requires delegated authentication (user context).")
        return 1

    # Build OData $filter for date range
    odata_filters: list[str] = []

    # Default to last year unless the user explicitly asked for everything
    after = args.after
    if not after and not args.fetch_all:
        after = "1y"
        logger.info(
            "📅 No time bound specified, defaulting to last year. Use --all to fetch everything."
        )

    if after:
        try:
            iso = parse_date_string(after)
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

    select_fields = ["toRecipients", "ccRecipients", "sentDateTime"]

    to_counter: Counter[str] = Counter()
    cc_counter: Counter[str] = Counter()
    names: dict[str, str] = {}

    if args.only != "received":
        query_params = MessagesRequestBuilder.MessagesRequestBuilderGetQueryParameters(
            select=select_fields,
            top=1000,
        )

        if odata_filters:
            query_params.filter = " and ".join(odata_filters)
            logger.info(f"📅 Date filter: {query_params.filter}")

        request_config = RequestConfiguration(query_parameters=query_params)

        logger.info("📨 Fetching sent messages...")
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

                for r in msg.cc_recipients or []:
                    if r.email_address and r.email_address.address:
                        addr = r.email_address.address.lower()
                        cc_counter[addr] += 1
                        if r.email_address.name:
                            names[addr] = r.email_address.name

            if result.odata_next_link:
                result = (
                    await context.graph_client.me.mail_folders.by_mail_folder_id(
                        "SentItems"
                    )
                    .messages.with_url(result.odata_next_link)
                    .get()
                )
            else:
                break

        logger.info(f"📊 Analysed {total:,} sent messages")
    top_n = args.top if args.top > 0 else None

    def _print_ranking(
        counter: Counter, names_dict: dict, title: str, emoji: str = "🏆"
    ) -> None:
        ranking = counter.most_common(top_n)
        if not ranking:
            return
        table = Table(title=f"{emoji} {title}", show_header=True, header_style="bold")
        table.add_column("#", justify="right", style="dim")
        table.add_column("Emails", justify="right")
        table.add_column("Contact")
        for rank, (addr, count) in enumerate(ranking, 1):
            name = names_dict.get(addr, "")
            display = f"{name} <{addr}>" if name else addr
            table.add_row(str(rank), str(count), display)
        Console().print(table)

    if args.only != "received":
        _print_ranking(
            to_counter, names, f"Top {args.top or len(to_counter)} — sent To"
        )
        _print_ranking(
            cc_counter, names, f"Top {args.top or len(cc_counter)} — sent CC"
        )

    if args.only == "sent":
        return 0

    # --- Received analysis ---
    user_email = ""
    if context.cached_user:
        user_email = (
            getattr(context.cached_user, "mail", None)
            or getattr(context.cached_user, "user_principal_name", None)
            or ""
        ).lower()

    if user_email:
        # Domains that are automated notification senders, not real contacts
        noise_domains = {
            "yammer.com",
            "engage.mail.microsoft",
            "sharepointonline.com",
            "docusign.net",
        }

        # KQL date bounds for $search (date portion only)
        kql_date_parts: list[str] = []
        if after:
            kql_date_parts.append(f"received>={parse_date_string(after).split('T')[0]}")
        if args.before:
            kql_date_parts.append(
                f"received<={parse_date_string(args.before).split('T')[0]}"
            )

        async def _search_senders(field: str) -> tuple[Counter[str], dict[str, str]]:
            """GET /me/messages?$search with KQL - server-side filtering, typed Message objects."""
            kql = '"' + " ".join([f"{field}:{user_email}"] + kql_date_parts) + '"'
            qp = MessagesRequestBuilder.MessagesRequestBuilderGetQueryParameters(
                select=["from"],
                top=1000,
                search=kql,
            )
            cfg = RequestConfiguration(query_parameters=qp)
            counter: Counter[str] = Counter()
            names_local: dict[str, str] = {}
            page = await context.graph_client.me.messages.get(request_configuration=cfg)
            while page:
                for msg in page.value or []:
                    if (
                        msg.from_
                        and msg.from_.email_address
                        and msg.from_.email_address.address
                    ):
                        sender = msg.from_.email_address.address.lower()
                        counter[sender] += 1
                        if msg.from_.email_address.name:
                            names_local[sender] = msg.from_.email_address.name
                if page.odata_next_link:
                    page = await context.graph_client.me.messages.with_url(
                        page.odata_next_link
                    ).get()
                else:
                    break
            return counter, names_local

        logger.info("📬 Fetching received messages...")
        recv_to_counter, recv_to_names = await _search_senders("to")
        recv_cc_counter, recv_cc_names = await _search_senders("cc")
        recv_names = {**recv_to_names, **recv_cc_names}
        logger.info(
            f"📊 Received: {sum(recv_to_counter.values()):,} as To, "
            f"{sum(recv_cc_counter.values()):,} as CC"
        )

        # Filter noise domains for display only (counts remain accurate)
        filtered_to = Counter(
            {
                k: v
                for k, v in recv_to_counter.items()
                if k.split("@", 1)[-1] not in noise_domains
            }
        )
        filtered_cc = Counter(
            {
                k: v
                for k, v in recv_cc_counter.items()
                if k.split("@", 1)[-1] not in noise_domains
            }
        )

        _print_ranking(
            filtered_to,
            recv_names,
            f"Top {args.top or len(filtered_to)} — received as To",
            "📥",
        )
        _print_ranking(
            filtered_cc,
            recv_names,
            f"Top {args.top or len(filtered_cc)} — received as CC",
            "📥",
        )
    else:
        logger.warning(
            "⚠️ Could not determine your email address; skipping received analysis."
        )

    if args.save:
        save_path = Path(args.save)
        save_path.parent.mkdir(parents=True, exist_ok=True)

        all_addrs = set(to_counter) | set(cc_counter)
        data = {
            "sent": [
                {
                    "rank": rank,
                    "email": addr,
                    "name": names.get(addr),
                    "to_count": to_counter.get(addr, 0),
                    "cc_count": cc_counter.get(addr, 0),
                }
                for rank, (addr, _) in enumerate(
                    sorted(all_addrs, key=lambda a: -to_counter.get(a, 0)), 1
                )
            ],
            "received": (
                [
                    {
                        "rank": rank,
                        "email": addr,
                        "name": recv_names.get(addr) if user_email else None,
                        "to_count": recv_to_counter.get(addr, 0) if user_email else 0,
                        "cc_count": recv_cc_counter.get(addr, 0) if user_email else 0,
                    }
                    for rank, (addr, _) in enumerate(
                        sorted(
                            set(recv_to_counter) | set(recv_cc_counter),
                            key=lambda a: -recv_to_counter.get(a, 0),
                        ),
                        1,
                    )
                ]
                if user_email
                else []
            ),
        }

        with open(save_path, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2, ensure_ascii=False)

        logger.info(f"💾 Full results saved to: {save_path.absolute()}")

    return 0
