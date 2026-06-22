# msgraphx/modules/aad/search.py

# Built-in imports
from __future__ import annotations

import argparse
import datetime
import json
from pathlib import Path

# External library imports
from loguru import logger
from msgraph.generated.groups.groups_request_builder import GroupsRequestBuilder
from msgraph.generated.users.users_request_builder import UsersRequestBuilder
from msgraph.generated.devices.devices_request_builder import DevicesRequestBuilder
from msgraph.generated.service_principals.service_principals_request_builder import (
    ServicePrincipalsRequestBuilder,
)
from msgraph.generated.applications.applications_request_builder import (
    ApplicationsRequestBuilder,
)

# Local library imports
from ...core.context import GraphContext
from ...utils import output, pagination
from ...utils.errors import handle_graph_errors

# Preset hunt queries for common privileged groups
HUNT_GROUPS = {
    "admins": ["admin", "administrator", "administrateur"],
    "domain": ["domain admin", "domain user", "domain controller", "domain"],
    "sql": ["sql", "database", "db"],
    "exchange": ["exchange", "mail", "email"],
    "backup": ["backup", "veeam", "commvault"],
    "security": ["security", "sec", "infosec"],
    "helpdesk": ["helpdesk", "support", "service desk"],
    "developers": ["dev", "developer", "development"],
    "privileged": ["privileged", "priv", "pam"],
    "azure-admins": ["Global Admin", "Privileged", "User Admin", "Security Admin"],
}


def matches_query(text: str, query: str, exact: bool = False) -> bool:
    if not text:
        return False
    text_lower = text.lower()
    query_lower = query.lower()
    if exact:
        return text_lower.startswith(query_lower)
    return query_lower in text_lower


def serialize_object(obj):
    if obj is None:
        return None
    if isinstance(obj, datetime.datetime):
        return obj.isoformat()
    if hasattr(obj, "__dict__"):
        result = {}
        for key, value in obj.__dict__.items():
            if key.startswith("_"):
                continue
            if isinstance(value, list):
                result[key] = [serialize_object(item) for item in value]
            elif isinstance(value, dict):
                result[key] = {k: serialize_object(v) for k, v in value.items()}
            elif hasattr(value, "__dict__"):
                result[key] = serialize_object(value)
            else:
                result[key] = value
        return result
    return obj


def save_results_to_json(
    tenant_id: str, search_type: str, query: str, results: list, output_dir: Path = None
):
    if not results:
        return None
    if output_dir is None:
        output_dir = Path.cwd()
    tenant_dir = output_dir / tenant_id
    tenant_dir.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.datetime.now(datetime.timezone.utc).strftime("%Y%m%d_%H%M%S")
    safe_query = "".join(c if c.isalnum() else "_" for c in query)
    filename = f"{search_type}_{safe_query}_{timestamp}.json"
    filepath = tenant_dir / filename
    serialized_results = [serialize_object(obj) for obj in results]
    data = {
        "tenant_id": tenant_id,
        "search_type": search_type,
        "query": query,
        "timestamp": datetime.datetime.now(datetime.timezone.utc).isoformat(),
        "count": len(results),
        "results": serialized_results,
    }
    with open(filepath, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)
    logger.success(f"Saved {len(results)} results to: {filepath}")
    return filepath


async def fetch_groups(
    context: GraphContext,
    query: str,
    contains: bool = False,
    synced_only: bool = False,
    odata_filter: str | None = None,
) -> list[dict]:
    """Return groups matching query as plain dicts.

    Raises on API error — callers are responsible for handling exceptions.
    """
    if odata_filter:
        logger.debug(f"Using custom OData filter: {odata_filter}")
        query_params = GroupsRequestBuilder.GroupsRequestBuilderGetQueryParameters(
            filter=odata_filter,
            top=999,
        )
        request_config = GroupsRequestBuilder.GroupsRequestBuilderGetRequestConfiguration(
            query_parameters=query_params
        )
        all_results = await pagination.collect_all(context.graph_client.groups, request_config)
    elif contains:
        logger.debug("Using client-side 'contains' filter (fetching all groups)")
        all_results = await pagination.collect_all(context.graph_client.groups, None)
        all_results = [
            g for g in all_results
            if matches_query(g.display_name, query) or matches_query(g.mail_nickname, query)
        ]
    else:
        query_params = GroupsRequestBuilder.GroupsRequestBuilderGetQueryParameters(
            filter=f"startswith(displayName,'{query}') or startswith(mailNickname,'{query}')",
            top=999,
        )
        request_config = GroupsRequestBuilder.GroupsRequestBuilderGetRequestConfiguration(
            query_parameters=query_params
        )
        all_results = await pagination.collect_all(context.graph_client.groups, request_config)

    if synced_only:
        all_results = [g for g in all_results if g.on_premises_sync_enabled]

    return [
        {
            "id": g.id,
            "display_name": g.display_name,
            "description": g.description,
            "mail": g.mail,
            "mail_nickname": g.mail_nickname,
            "security_enabled": g.security_enabled,
            "mail_enabled": g.mail_enabled,
            "on_premises_sync_enabled": g.on_premises_sync_enabled,
            "visibility": g.visibility,
            "group_types": g.group_types,
        }
        for g in all_results
    ]


async def fetch_users(
    context: GraphContext,
    query: str,
    odata_filter: str | None = None,
) -> list[dict]:
    """Return users matching query as plain dicts.

    Raises on API error — callers are responsible for handling exceptions.
    """
    api_filter = (
        odata_filter
        if odata_filter
        else f"startswith(displayName,'{query}') or startswith(userPrincipalName,'{query}') or startswith(mail,'{query}')"
    )
    if odata_filter:
        logger.debug(f"Using custom OData filter: {odata_filter}")
    query_params = UsersRequestBuilder.UsersRequestBuilderGetQueryParameters(
        filter=api_filter,
        top=999,
    )
    request_config = UsersRequestBuilder.UsersRequestBuilderGetRequestConfiguration(
        query_parameters=query_params
    )
    all_results = await pagination.collect_all(context.graph_client.users, request_config)

    return [
        {
            "id": u.id,
            "display_name": u.display_name,
            "user_principal_name": u.user_principal_name,
            "mail": u.mail,
            "account_enabled": u.account_enabled,
            "job_title": u.job_title,
            "department": u.department,
        }
        for u in all_results
    ]


async def fetch_devices(
    context: GraphContext,
    query: str,
    odata_filter: str | None = None,
) -> list[dict]:
    """Return devices matching query as plain dicts.

    Raises on API error — callers are responsible for handling exceptions.
    """
    api_filter = odata_filter if odata_filter else f"startswith(displayName,'{query}')"
    if odata_filter:
        logger.debug(f"Using custom OData filter: {odata_filter}")
    query_params = DevicesRequestBuilder.DevicesRequestBuilderGetQueryParameters(
        filter=api_filter,
        top=999,
    )
    request_config = DevicesRequestBuilder.DevicesRequestBuilderGetRequestConfiguration(
        query_parameters=query_params
    )
    all_results = await pagination.collect_all(context.graph_client.devices, request_config)

    return [
        {
            "id": d.id,
            "display_name": d.display_name,
            "account_enabled": d.account_enabled,
            "operating_system": d.operating_system,
            "operating_system_version": d.operating_system_version,
            "trust_type": d.trust_type,
            "on_premises_sync_enabled": d.on_premises_sync_enabled,
        }
        for d in all_results
    ]


async def fetch_service_principals(
    context: GraphContext,
    query: str,
    odata_filter: str | None = None,
) -> list[dict]:
    """Return service principals matching query as plain dicts.

    Raises on API error — callers are responsible for handling exceptions.
    """
    api_filter = odata_filter if odata_filter else f"startswith(displayName,'{query}')"
    if odata_filter:
        logger.debug(f"Using custom OData filter: {odata_filter}")
    query_params = ServicePrincipalsRequestBuilder.ServicePrincipalsRequestBuilderGetQueryParameters(
        filter=api_filter,
        top=999,
    )
    request_config = ServicePrincipalsRequestBuilder.ServicePrincipalsRequestBuilderGetRequestConfiguration(
        query_parameters=query_params
    )
    all_results = await pagination.collect_all(
        context.graph_client.service_principals, request_config
    )

    return [
        {
            "id": sp.id,
            "display_name": sp.display_name,
            "app_id": sp.app_id,
            "account_enabled": sp.account_enabled,
            "service_principal_type": sp.service_principal_type,
        }
        for sp in all_results
    ]


async def fetch_applications(
    context: GraphContext,
    query: str,
    odata_filter: str | None = None,
) -> list[dict]:
    """Return applications matching query as plain dicts.

    Raises on API error — callers are responsible for handling exceptions.
    """
    api_filter = odata_filter if odata_filter else f"startswith(displayName,'{query}')"
    if odata_filter:
        logger.debug(f"Using custom OData filter: {odata_filter}")
    query_params = ApplicationsRequestBuilder.ApplicationsRequestBuilderGetQueryParameters(
        filter=api_filter,
        top=999,
    )
    request_config = ApplicationsRequestBuilder.ApplicationsRequestBuilderGetRequestConfiguration(
        query_parameters=query_params
    )
    all_results = await pagination.collect_all(
        context.graph_client.applications, request_config
    )

    return [
        {
            "id": app.id,
            "display_name": app.display_name,
            "app_id": app.app_id,
            "sign_in_audience": app.sign_in_audience,
            "created_date_time": app.created_date_time.isoformat() if app.created_date_time else None,
        }
        for app in all_results
    ]


def add_arguments(parser: "argparse.ArgumentParser"):
    parser.add_argument(
        "query",
        nargs="?",
        default=None,
        help="Search query (e.g., 'admin', 'SQL', 'Domain'). Use with --type to specify what to search.",
    )

    parser.add_argument(
        "-t",
        "--type",
        type=str,
        choices=[
            "groups",
            "users",
            "devices",
            "computers",
            "service-principals",
            "applications",
            "all",
        ],
        default="groups",
        help="Type of Azure AD object to search. Default: groups. Use 'all' to search everything.",
    )

    parser.add_argument(
        "--hunt",
        choices=HUNT_GROUPS.keys(),
        help="Preset hunt queries for common privileged groups (admins, domain, sql, exchange, etc.)",
    )

    parser.add_argument(
        "--enabled-only",
        action="store_true",
        help="Only show enabled objects (excludes disabled accounts/groups).",
    )

    parser.add_argument(
        "--contains",
        action="store_true",
        help="Use client-side 'contains' matching instead of API 'startswith' (slower but more flexible).",
    )

    parser.add_argument(
        "--synced-only",
        action="store_true",
        help="Only show groups/users synced from on-premises Active Directory.",
    )

    parser.add_argument(
        "--filter",
        dest="odata_filter",
        type=str,
        default=None,
        metavar="EXPR",
        help="Raw OData \\$filter expression passed directly to the API, e.g. \"accountEnabled eq true\". Overrides the default startswith filter.",
    )

    parser.add_argument(
        "--save-dir",
        dest="save_dir",
        type=str,
        default=None,
        help="Save results as JSON files under this directory. Structure: {dir}/{tenant_id}/{type}_{query}_{timestamp}.json",
    )


def _log_groups(results: list[dict]) -> None:
    for g in results:
        group_type = "Security" if g["security_enabled"] else "Distribution"
        if g["group_types"] and "Unified" in g["group_types"]:
            group_type = "Microsoft 365"
        sync_indicator = "Synced" if g["on_premises_sync_enabled"] else ""
        logger.success(f"{g['display_name']} | Type: {group_type} {sync_indicator} | ID: {g['id']}")
        if g["description"]:
            logger.info(f"   Description: {g['description']}")
        if g["mail"]:
            logger.info(f"   Email: {g['mail']}")


def _log_users(results: list[dict]) -> None:
    for u in results:
        status = "Enabled" if u["account_enabled"] else "Disabled"
        logger.success(f"{u['display_name']} ({u['user_principal_name']}) | {status} | ID: {u['id']}")
        if u["job_title"]:
            logger.info(f"   Title: {u['job_title']}")
        if u["department"]:
            logger.info(f"   Department: {u['department']}")


def _log_devices(results: list[dict]) -> None:
    for d in results:
        status = "Enabled" if d["account_enabled"] else "Disabled"
        os_info = (
            f"{d['operating_system']} {d['operating_system_version']}"
            if d["operating_system"]
            else "Unknown OS"
        )
        logger.success(f"{d['display_name']} | {os_info} | {status} | ID: {d['id']}")
        if d["trust_type"]:
            logger.info(f"   Trust Type: {d['trust_type']}")


def _log_service_principals(results: list[dict]) -> None:
    for sp in results:
        status = "Enabled" if sp["account_enabled"] else "Disabled"
        logger.success(f"{sp['display_name']} | App ID: {sp['app_id']} | {status} | ID: {sp['id']}")
        if sp["service_principal_type"]:
            logger.info(f"   Type: {sp['service_principal_type']}")


def _log_applications(results: list[dict]) -> None:
    for app in results:
        logger.success(f"{app['display_name']} | App ID: {app['app_id']} | Audience: {app['sign_in_audience']} | ID: {app['id']}")
        if app["created_date_time"]:
            logger.info(f"   Created: {app['created_date_time']}")


@handle_graph_errors
async def run_with_arguments(
    context: "GraphContext", args: "argparse.Namespace"
) -> int:
    if args.hunt:
        queries = HUNT_GROUPS[args.hunt]
        logger.info(f"Hunt mode: {args.hunt}")
        logger.info(f"Keywords: {', '.join(queries)}")
    elif args.query:
        queries = [args.query]
    else:
        logger.error("Please provide a search query or use --hunt")
        return 1

    search_types = []
    if args.type == "all":
        search_types = ["groups", "users", "devices", "service-principals", "applications"]
    elif args.type == "computers":
        search_types = ["devices"]
    else:
        search_types = [args.type]

    tenant_id = getattr(args, "tenant_id", None)
    save_dir = getattr(args, "save_dir", None)
    save_json = save_dir is not None
    output_dir = Path(save_dir) if save_dir else Path.cwd()
    odata_filter: str | None = getattr(args, "odata_filter", None)

    if save_json and not tenant_id:
        logger.warning("Tenant ID not available, using 'unknown' for directory name")
        tenant_id = "unknown"

    total_found = 0
    structured_results: dict[str, list] = {}

    for query in queries:
        for search_type in search_types:
            try:
                if search_type == "groups":
                    results = await fetch_groups(
                        context, query,
                        contains=args.contains,
                        synced_only=args.synced_only,
                        odata_filter=odata_filter,
                    )
                    _log_groups(results)
                elif search_type == "users":
                    results = await fetch_users(context, query, odata_filter=odata_filter)
                    _log_users(results)
                elif search_type in ("devices", "computers"):
                    results = await fetch_devices(context, query, odata_filter=odata_filter)
                    _log_devices(results)
                elif search_type == "service-principals":
                    results = await fetch_service_principals(context, query, odata_filter=odata_filter)
                    _log_service_principals(results)
                elif search_type == "applications":
                    results = await fetch_applications(context, query, odata_filter=odata_filter)
                    _log_applications(results)
                else:
                    results = []

            except Exception as exc:
                logger.error(f"Failed to search {search_type}: {exc}")
                results = []

            if not results:
                logger.info(f"No {search_type} found for query: {query!r}")

            if results and save_json and tenant_id:
                save_results_to_json(tenant_id, search_type, query, results, output_dir)

            total_found += len(results)
            structured_results.setdefault(search_type, []).extend(results)

            if len(queries) > 1 or len(search_types) > 1:
                logger.info("─" * 80)

    logger.info(f"Total objects found: {total_found}")

    if context.ndjson_output:
        for items in structured_results.values():
            for item in items:
                output.print_ndjson_item(item)
    elif context.json_output:
        output.print_json(structured_results)

    return 0
