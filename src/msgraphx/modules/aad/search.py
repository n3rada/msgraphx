# msgraphx/modules/aad/search.py

# Built-in imports
from typing import TYPE_CHECKING
import json
from pathlib import Path
from datetime import datetime

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
from msgraphx.utils.pagination import collect_all


if TYPE_CHECKING:
    import argparse
    from msgraphx.core.context import GraphContext


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
    """Check if text matches query (case-insensitive)."""
    if not text:
        return False
    text_lower = text.lower()
    query_lower = query.lower()
    if exact:
        return text_lower.startswith(query_lower)
    return query_lower in text_lower


def serialize_object(obj):
    """Convert Graph API objects to JSON-serializable format."""
    if obj is None:
        return None

    # Handle datetime objects
    if isinstance(obj, datetime):
        return obj.isoformat()

    # Handle objects with __dict__ (common in msgraph SDK)
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
    """Save search results to JSON file."""
    if not results:
        return None

    # Create output directory structure
    if output_dir is None:
        output_dir = Path.cwd()

    tenant_dir = output_dir / tenant_id
    tenant_dir.mkdir(parents=True, exist_ok=True)

    # Generate filename with timestamp
    timestamp = datetime.utcnow().strftime("%Y%m%d_%H%M%S")
    safe_query = "".join(c if c.isalnum() else "_" for c in query)
    filename = f"{search_type}_{safe_query}_{timestamp}.json"
    filepath = tenant_dir / filename

    # Serialize results
    serialized_results = [serialize_object(obj) for obj in results]

    # Save to file
    data = {
        "tenant_id": tenant_id,
        "search_type": search_type,
        "query": query,
        "timestamp": datetime.utcnow().isoformat(),
        "count": len(results),
        "results": serialized_results,
    }

    with open(filepath, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)

    logger.success(f"💾 Saved {len(results)} results to: {filepath}")
    return filepath


async def search_groups(
    graph_client: "GraphServiceClient",
    query: str,
    save_json: bool = False,
    tenant_id: str = None,
    output_dir: Path = None,
    contains: bool = False,
    show_synced_only: bool = False,
) -> int:
    """Search for groups matching the query."""
    logger.info(f"🔍 Searching for groups matching: {query}")

    count = 0
    all_results = []

    try:
        # For contains search, get all groups and filter client-side
        if contains:
            logger.debug("🔎 Using client-side 'contains' filter (fetching all groups)")
            all_results = await collect_all(graph_client.groups, None)
            # Filter client-side
            all_results = [
                g
                for g in all_results
                if matches_query(g.display_name, query)
                or matches_query(g.mail_nickname, query)
            ]
        else:
            # API-side startswith filter
            query_params = GroupsRequestBuilder.GroupsRequestBuilderGetQueryParameters(
                filter=f"startswith(displayName,'{query}') or startswith(mailNickname,'{query}')",
                top=999,
            )
            request_config = (
                GroupsRequestBuilder.GroupsRequestBuilderGetRequestConfiguration(
                    query_parameters=query_params
                )
            )
            all_results = await collect_all(graph_client.groups, request_config)

        # Filter for synced groups if requested
        if show_synced_only:
            all_results = [g for g in all_results if g.on_premises_sync_enabled]
            logger.debug(
                f"🔄 Filtered to {len(all_results)} synced from on-premises AD"
            )

        if all_results:
            for group in all_results:
                group_type = "Security" if group.security_enabled else "Distribution"
                if group.group_types and "Unified" in group.group_types:
                    group_type = "Microsoft 365"

                sync_indicator = "🔄 Synced" if group.on_premises_sync_enabled else ""
                logger.success(
                    f"👥 {group.display_name} | Type: {group_type} {sync_indicator} | ID: {group.id}"
                )
                if group.description:
                    logger.info(f"   Description: {group.description}")
                if group.mail:
                    logger.info(f"   Email: {group.mail}")
                count += 1

        if count == 0:
            logger.info("📭 No groups found")
        elif save_json and tenant_id:
            save_results_to_json(tenant_id, "groups", query, all_results, output_dir)

    except Exception as e:
        logger.error(f"❌ Failed to search groups: {e}")

    return count


async def search_users(
    graph_client: "GraphServiceClient",
    query: str,
    save_json: bool = False,
    tenant_id: str = None,
    output_dir: Path = None,
) -> int:
    """Search for users matching the query."""
    logger.info(f"🔍 Searching for users matching: {query}")

    count = 0
    all_results = []

    try:
        # Build filter for display name or userPrincipalName contains query
        query_params = UsersRequestBuilder.UsersRequestBuilderGetQueryParameters(
            filter=f"startswith(displayName,'{query}') or startswith(userPrincipalName,'{query}') or startswith(mail,'{query}')",
            top=999,
        )
        request_config = UsersRequestBuilder.UsersRequestBuilderGetRequestConfiguration(
            query_parameters=query_params
        )

        # Collect all pages of results
        all_results = await collect_all(graph_client.users, request_config)

        if all_results:
            for user in all_results:
                status = "✅ Enabled" if user.account_enabled else "🔒 Disabled"
                logger.success(
                    f"👤 {user.display_name} ({user.user_principal_name}) | {status} | ID: {user.id}"
                )
                if user.job_title:
                    logger.info(f"   Title: {user.job_title}")
                if user.department:
                    logger.info(f"   Department: {user.department}")
                count += 1

        if count == 0:
            logger.info("📭 No users found")
        elif save_json and tenant_id:
            save_results_to_json(tenant_id, "users", query, all_results, output_dir)

    except Exception as e:
        logger.error(f"❌ Failed to search users: {e}")

    return count


async def search_devices(
    graph_client: "GraphServiceClient",
    query: str,
    save_json: bool = False,
    tenant_id: str = None,
    output_dir: Path = None,
) -> int:
    """Search for devices/computers matching the query."""
    logger.info(f"🔍 Searching for devices matching: {query}")

    count = 0
    all_results = []

    try:
        # Build filter for display name contains query
        query_params = DevicesRequestBuilder.DevicesRequestBuilderGetQueryParameters(
            filter=f"startswith(displayName,'{query}')",
            top=999,
        )
        request_config = (
            DevicesRequestBuilder.DevicesRequestBuilderGetRequestConfiguration(
                query_parameters=query_params
            )
        )

        # Collect all pages of results
        all_results = await collect_all(graph_client.devices, request_config)

        if all_results:
            for device in all_results:
                status = "✅ Enabled" if device.account_enabled else "🔒 Disabled"
                os_info = (
                    f"{device.operating_system} {device.operating_system_version}"
                    if device.operating_system
                    else "Unknown OS"
                )

                logger.success(
                    f"💻 {device.display_name} | {os_info} | {status} | ID: {device.id}"
                )
                if device.trust_type:
                    logger.info(f"   Trust Type: {device.trust_type}")
                count += 1

        if count == 0:
            logger.info("📭 No devices found")
        elif save_json and tenant_id:
            save_results_to_json(tenant_id, "devices", query, all_results, output_dir)

    except Exception as e:
        logger.error(f"❌ Failed to search devices: {e}")

    return count


async def search_service_principals(
    graph_client: "GraphServiceClient",
    query: str,
    save_json: bool = False,
    tenant_id: str = None,
    output_dir: Path = None,
) -> int:
    """Search for service principals matching the query."""
    logger.info(f"🔍 Searching for service principals matching: {query}")

    count = 0
    all_results = []

    try:
        query_params = ServicePrincipalsRequestBuilder.ServicePrincipalsRequestBuilderGetQueryParameters(
            filter=f"startswith(displayName,'{query}')",
            top=999,
        )
        request_config = ServicePrincipalsRequestBuilder.ServicePrincipalsRequestBuilderGetRequestConfiguration(
            query_parameters=query_params
        )

        # Collect all pages of results
        all_results = await collect_all(graph_client.service_principals, request_config)

        if all_results:
            for sp in all_results:
                status = "✅ Enabled" if sp.account_enabled else "🔒 Disabled"
                logger.success(
                    f"🤖 {sp.display_name} | App ID: {sp.app_id} | {status} | ID: {sp.id}"
                )
                if sp.service_principal_type:
                    logger.info(f"   Type: {sp.service_principal_type}")
                count += 1

        if count == 0:
            logger.info("📭 No service principals found")
        elif save_json and tenant_id:
            save_results_to_json(
                tenant_id, "service_principals", query, all_results, output_dir
            )

    except Exception as e:
        logger.error(f"❌ Failed to search service principals: {e}")

    return count


async def search_applications(
    graph_client: "GraphServiceClient",
    query: str,
    save_json: bool = False,
    tenant_id: str = None,
    output_dir: Path = None,
) -> int:
    """Search for applications matching the query."""
    logger.info(f"🔍 Searching for applications matching: {query}")

    count = 0
    all_results = []

    try:
        query_params = (
            ApplicationsRequestBuilder.ApplicationsRequestBuilderGetQueryParameters(
                filter=f"startswith(displayName,'{query}')",
                top=999,
            )
        )
        request_config = ApplicationsRequestBuilder.ApplicationsRequestBuilderGetRequestConfiguration(
            query_parameters=query_params
        )

        # Collect all pages of results
        all_results = await collect_all(graph_client.applications, request_config)

        if all_results:
            for app in all_results:
                logger.success(
                    f"📱 {app.display_name} | App ID: {app.app_id} | Audience: {app.sign_in_audience} | ID: {app.id}"
                )
                if app.created_date_time:
                    logger.info(f"   Created: {app.created_date_time}")
                count += 1

        if count == 0:
            logger.info("📭 No applications found")
        elif save_json and tenant_id:
            save_results_to_json(
                tenant_id, "applications", query, all_results, output_dir
            )

    except Exception as e:
        logger.error(f"❌ Failed to search applications: {e}")

    return count


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
        "-o",
        "--output",
        type=str,
        default=None,
        help="Output directory for JSON results. Defaults to current directory. Results saved as: {output}/{tenant_id}/{type}_{query}_{timestamp}.json",
    )


async def run_with_arguments(
    context: "GraphContext", args: "argparse.Namespace"
) -> int:

    # Determine the search query
    if args.hunt:
        # Use hunt preset
        hunt_keywords = HUNT_GROUPS[args.hunt]
        logger.info(f"🎯 Hunt mode: {args.hunt}")
        logger.info(f"🔍 Keywords: {', '.join(hunt_keywords)}")
        queries = hunt_keywords
    elif args.query:
        queries = [args.query]
    else:
        logger.error("❌ Please provide a search query or use --hunt")
        return 1

    # Determine which types to search
    search_types = []
    if args.type == "all":
        search_types = [
            "groups",
            "users",
            "devices",
            "service-principals",
            "applications",
        ]
    elif args.type == "computers":
        search_types = ["devices"]  # Alias for devices
    else:
        search_types = [args.type]

    # Get tenant ID and output settings
    tenant_id = getattr(args, "tenant_id", None)
    save_json = args.output is not None
    output_dir = Path(args.output) if args.output else Path.cwd()

    if save_json and not tenant_id:
        logger.warning("⚠️ Tenant ID not available, using 'unknown' for directory name")
        tenant_id = "unknown"

    total_found = 0

    # Execute searches
    for query in queries:
        for search_type in search_types:
            if search_type == "groups":
                total_found += await search_groups(
                    context.graph_client,
                    query,
                    save_json,
                    tenant_id,
                    output_dir,
                    contains=args.contains,
                    show_synced_only=args.synced_only,
                )
            elif search_type == "users":
                total_found += await search_users(
                    context.graph_client, query, save_json, tenant_id, output_dir
                )
            elif search_type in ["devices", "computers"]:
                total_found += await search_devices(
                    context.graph_client, query, save_json, tenant_id, output_dir
                )
            elif search_type == "service-principals":
                total_found += await search_service_principals(
                    context.graph_client, query, save_json, tenant_id, output_dir
                )
            elif search_type == "applications":
                total_found += await search_applications(
                    context.graph_client, query, save_json, tenant_id, output_dir
                )

            # Add separator between different searches
            if len(queries) > 1 or len(search_types) > 1:
                logger.info("─" * 80)

    logger.info(f"📊 Total objects found: {total_found}")
    return 0
