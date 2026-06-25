# msgraphx/modules/outlook/contacts.py
#
# Build a communication graph from your mail: who do you exchange with most?
# Required delegated permission: Mail.ReadBasic (or Mail.Read)

# Built-in imports
from __future__ import annotations

import argparse
import contextlib
import importlib.resources
import json
from collections import Counter
from pathlib import Path


def _load_config(filename: str) -> frozenset[str]:
    return frozenset(
        line.strip()
        for line in importlib.resources.files("msgraphx")
        .joinpath(f"config/{filename}")
        .read_text(encoding="utf-8")
        .splitlines()
        if line.strip() and not line.startswith("#")
    )


_NOISE_DOMAINS: frozenset[str] = _load_config("noise_domains.txt")
_NOISE_LOCALS: frozenset[str] = _load_config("noise_locals.txt")

# External library imports
from kiota_abstractions.base_request_configuration import RequestConfiguration
from loguru import logger
from msgraph.generated.users.item.mail_folders.item.messages.messages_request_builder import (
    MessagesRequestBuilder,
)
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


async def fetch(
    context: GraphContext,
    after: str | None = None,
    before: str | None = None,
    only: str | None = None,
    top: int | None = None,
) -> dict:
    """Analyse mail to build a communication graph. Returns sent/received contact counts.

    `after` / `before` are ISO 8601 strings. `only` is 'sent', 'received', or None (both).
    Returns {'sent': [...], 'received': [...]}.
    Raises on API error. Callers are responsible for handling exceptions.
    """
    select_fields = ["toRecipients", "ccRecipients", "sentDateTime"]

    to_counter: Counter[str] = Counter()
    cc_counter: Counter[str] = Counter()
    names: dict[str, str] = {}
    recv_to_counter: Counter[str] = Counter()
    recv_cc_counter: Counter[str] = Counter()
    recv_names: dict[str, str] = {}

    if only != "received":
        odata_filters: list[str] = []
        if after:
            odata_filters.append(f"sentDateTime ge {after}")
        if before:
            odata_filters.append(f"sentDateTime le {before}")

        query_params = MessagesRequestBuilder.MessagesRequestBuilderGetQueryParameters(
            select=select_fields,
            top=1000,
        )
        if odata_filters:
            query_params.filter = " and ".join(odata_filters)

        request_config = RequestConfiguration(query_parameters=query_params)
        sent_builder = context.graph_client.me.mail_folders.by_mail_folder_id(
            "SentItems"
        ).messages

        async for msg in GraphPaginator(sent_builder, request_config):
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

    if only != "sent":
        user_email = ""
        if context.cached_user:
            user_email = (
                getattr(context.cached_user, "mail", None)
                or getattr(context.cached_user, "user_principal_name", None)
                or ""
            ).lower()

        if user_email:
            def _is_noisy(addr: str) -> bool:
                local, _, domain = addr.partition("@")
                normalised = local.replace("-", "").replace("_", "").lower()
                return (
                    any(domain == nd or domain.endswith("." + nd) for nd in _NOISE_DOMAINS)
                    or normalised in _NOISE_LOCALS
                    or "noreply" in normalised
                )

            recv_filter_parts: list[str] = []
            if after:
                recv_filter_parts.append(f"receivedDateTime ge {after}")
            if before:
                recv_filter_parts.append(f"receivedDateTime le {before}")

            recv_qp = MessagesRequestBuilder.MessagesRequestBuilderGetQueryParameters(
                select=["from", "toRecipients", "ccRecipients"],
                top=50,
            )
            if recv_filter_parts:
                recv_qp.filter = " and ".join(recv_filter_parts)

            recv_cfg = RequestConfiguration(query_parameters=recv_qp)
            inbox_builder = context.graph_client.me.mail_folders.by_mail_folder_id(
                "Inbox"
            ).messages

            async for msg in GraphPaginator(inbox_builder, recv_cfg):
                if not (
                    msg.from_
                    and msg.from_.email_address
                    and msg.from_.email_address.address
                ):
                    continue

                sender = msg.from_.email_address.address.lower()
                if _is_noisy(sender):
                    continue

                if msg.from_.email_address.name:
                    recv_names[sender] = msg.from_.email_address.name

                to_addrs = {
                    r.email_address.address.lower()
                    for r in (msg.to_recipients or [])
                    if r.email_address and r.email_address.address
                }
                cc_addrs = {
                    r.email_address.address.lower()
                    for r in (msg.cc_recipients or [])
                    if r.email_address and r.email_address.address
                }

                if user_email in to_addrs:
                    recv_to_counter[sender] += 1
                elif user_email in cc_addrs:
                    recv_cc_counter[sender] += 1

    all_sent_addrs = set(to_counter) | set(cc_counter)
    return {
        "sent": [
            {
                "rank": rank,
                "email": addr,
                "name": names.get(addr),
                "to_count": to_counter.get(addr, 0),
                "cc_count": cc_counter.get(addr, 0),
            }
            for rank, (addr, _) in enumerate(
                sorted(all_sent_addrs, key=lambda a: -to_counter.get(a, 0)),
                1,
            )
        ][:top],
        "received": [
            {
                "rank": rank,
                "email": addr,
                "name": recv_names.get(addr),
                "to_count": recv_to_counter.get(addr, 0),
                "cc_count": recv_cc_counter.get(addr, 0),
            }
            for rank, (addr, _) in enumerate(
                sorted(
                    set(recv_to_counter) | set(recv_cc_counter),
                    key=lambda a: -recv_to_counter.get(a, 0),
                ),
                1,
            )
        ][:top],
    }


def add_arguments(parser: "argparse.ArgumentParser"):
    parser.set_defaults(uses_time_bounds=True)
    parser.add_argument(
        "--top",
        "-n",
        type=int,
        default=20,
        metavar="N",
        help="Number of top contacts to display (default: 20). Use 0 for all.",
    )

    parser.add_argument(
        "--only",
        choices=["sent", "received"],
        default=None,
        metavar="{sent,received}",
        help="Restrict analysis to sent or received mail only (default: both).",
    )


@handle_graph_errors
@require_scopes("Mail.ReadBasic")
async def run_with_arguments(
    context: "GraphContext", args: "argparse.Namespace"
) -> int:
    if context.is_app_only:
        logger.error("This module requires delegated authentication (user context).")
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

    select_fields = ["toRecipients", "ccRecipients", "sentDateTime"]

    to_counter: Counter[str] = Counter()
    cc_counter: Counter[str] = Counter()
    names: dict[str, str] = {}

    top_n = args.top if args.top > 0 else None

    def _build_table(counter: Counter, names_dict: dict, title: str) -> Table:
        table = Table(title=title, show_header=True, header_style="bold")
        table.add_column("#", justify="right", style="dim")
        table.add_column("Emails", justify="right")
        table.add_column("Contact")
        for rank, (addr, count) in enumerate(counter.most_common(top_n), 1):
            name = names_dict.get(addr, "")
            display = f"{name} <{addr}>" if name else addr
            table.add_row(str(rank), str(count), display)
        return table

    if args.only != "received":
        query_params = MessagesRequestBuilder.MessagesRequestBuilderGetQueryParameters(
            select=select_fields,
            top=1000,
        )

        if odata_filters:
            query_params.filter = " and ".join(odata_filters)
            logger.info(f"Date filter: {query_params.filter}")

        request_config = RequestConfiguration(query_parameters=query_params)

        logger.info("Fetching sent messages")
        total = 0

        sent_builder = context.graph_client.me.mail_folders.by_mail_folder_id(
            "SentItems"
        ).messages

        live_ctx = contextlib.nullcontext() if context.json_output else Live(console=console, refresh_per_second=4)
        with live_ctx as live:
            async for msg in GraphPaginator(sent_builder, request_config):
                total += 1
                if live is not None:
                    live.update(
                        Group(
                            _build_table(
                                to_counter,
                                names,
                                f"Top {args.top or len(to_counter)} : sent To ({total:,} scanned)",
                            ),
                            _build_table(
                                cc_counter,
                                names,
                                f"Top {args.top or len(cc_counter)} : sent CC",
                            ),
                        )
                    )

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

        logger.info(f"Analysed {total:,} sent messages")

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

    # Initialised here so the --save block can reference them even when
    # user_email is empty or received analysis is skipped.
    recv_to_counter: Counter[str] = Counter()
    recv_cc_counter: Counter[str] = Counter()
    recv_names: dict[str, str] = {}

    if user_email:

        def _is_noisy(addr: str) -> bool:
            local, _, domain = addr.partition("@")
            normalised = local.replace("-", "").replace("_", "").lower()
            return (
                any(domain == nd or domain.endswith("." + nd) for nd in _NOISE_DOMAINS)
                or normalised in _NOISE_LOCALS
                or "noreply" in normalised
            )

        recv_filter_parts: list[str] = []
        if args.after:
            try:
                recv_filter_parts.append(
                    f"receivedDateTime ge {parse_date_string(args.after)}"
                )
            except ValueError as e:
                logger.error(str(e))
                return 1
        if args.before:
            try:
                recv_filter_parts.append(
                    f"receivedDateTime le {parse_date_string(args.before)}"
                )
            except ValueError as e:
                logger.error(str(e))
                return 1

        recv_qp = MessagesRequestBuilder.MessagesRequestBuilderGetQueryParameters(
            select=["from", "toRecipients", "ccRecipients"],
            top=50,
        )
        if recv_filter_parts:
            recv_qp.filter = " and ".join(recv_filter_parts)
            logger.info(f"Inbox filter: {recv_qp.filter}")

        recv_cfg = RequestConfiguration(query_parameters=recv_qp)
        inbox_builder = context.graph_client.me.mail_folders.by_mail_folder_id(
            "Inbox"
        ).messages

        logger.info("Fetching received messages from Inbox")

        total_recv = 0
        recv_live_ctx = contextlib.nullcontext() if context.json_output else Live(console=console, refresh_per_second=4)
        with recv_live_ctx as live:
            async for msg in GraphPaginator(inbox_builder, recv_cfg):
                total_recv += 1
                if live is not None:
                    live.update(
                        Group(
                            _build_table(
                                recv_to_counter,
                                recv_names,
                                f"Top {args.top or len(recv_to_counter)} : received as To ({total_recv:,} scanned)",
                            ),
                            _build_table(
                                recv_cc_counter,
                                recv_names,
                                f"Top {args.top or len(recv_cc_counter)} : received as CC",
                            ),
                        )
                    )

                if not (
                    msg.from_
                    and msg.from_.email_address
                    and msg.from_.email_address.address
                ):
                    continue

                sender = msg.from_.email_address.address.lower()
                if _is_noisy(sender):
                    continue

                if msg.from_.email_address.name:
                    recv_names[sender] = msg.from_.email_address.name

                to_addrs = {
                    r.email_address.address.lower()
                    for r in (msg.to_recipients or [])
                    if r.email_address and r.email_address.address
                }
                cc_addrs = {
                    r.email_address.address.lower()
                    for r in (msg.cc_recipients or [])
                    if r.email_address and r.email_address.address
                }

                if user_email in to_addrs:
                    recv_to_counter[sender] += 1
                elif user_email in cc_addrs:
                    recv_cc_counter[sender] += 1

        logger.info(
            f"Received: {sum(recv_to_counter.values()):,} as To, "
            f"{sum(recv_cc_counter.values()):,} as CC"
        )
    else:
        logger.warning(
            "Could not determine your email address; skipping received analysis."
        )

    if args.save or context.json_output:
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

        if args.save:
            save_path = Path(args.save)
            save_path.parent.mkdir(parents=True, exist_ok=True)
            with open(save_path, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=2, ensure_ascii=False)
            logger.info(f"Full results saved to: {save_path.absolute()}")

        if context.json_output:
            output.print_json(data)

    return 0
