# msgraphx/core/graph_search.py

# Built-in imports
from __future__ import annotations

from collections.abc import AsyncGenerator
from dataclasses import dataclass

# External library imports
from loguru import logger
from msgraph.graph_service_client import GraphServiceClient
from msgraph.generated.models.entity import Entity
from msgraph.generated.models.search_request import SearchRequest
from msgraph.generated.models.search_query import SearchQuery
from msgraph.generated.models.sort_property import SortProperty
from msgraph.generated.search.query.query_post_request_body import QueryPostRequestBody
from msgraph.generated.models.entity_type import EntityType
from msgraph.generated.models.share_point_one_drive_options import SharePointOneDriveOptions

# Local library imports
from ..utils.errors import raise_if_forbidden


@dataclass
class SearchOptions:
    """Settings for a Microsoft Search API query.

    Attributes:
        query_string: KQL query string; `*` matches everything.
        sort_by: Property to sort on, or `None` for relevance order. Ignored
            for message and event searches, which do not support sorting.
        descending: Sort in descending order.
        fields: Properties to return, or `None` for the defaults.
        page_size: Results requested per page.
        region: Search region; required with application permissions.
        drive_id: Restrict results to this drive.
        max_pages: Maximum pages to fetch, or `None` for no cap (Exchange
            limits `from + size` to 1000).
        share_point_one_drive_options: SharePoint and OneDrive search scope
            options.
        site_path: Restrict results to this site URL via KQL `Path:`.
    """

    query_string: str = "*"
    sort_by: str | None = "createdDateTime"
    descending: bool = True
    fields: list[str] | None = None
    page_size: int = 500
    region: str | None = None  # Required for application permissions
    drive_id: str | None = None  # Scope search to specific drive
    max_pages: int | None = (
        None  # Hard cap on pages (e.g. Exchange caps from+size≤1000)
    )
    share_point_one_drive_options: SharePointOneDriveOptions | None = None
    site_path: str | None = None  # Scope to a site via KQL Path:"<url>"


async def search_entities(
    client: GraphServiceClient,
    entity_types: list[EntityType],
    options: SearchOptions = SearchOptions(),
) -> AsyncGenerator[Entity, None]:
    """Search Microsoft 365 content with the Graph Search API, across all pages.

    Sends `POST /search/query` one page at a time until the server reports
    no more results, a page is empty, or `options.max_pages` is reached. If
    sorting is requested for message or event searches, a warning is logged
    and sorting is disabled. Errors other than 403 are logged and end the
    search. See
    https://learn.microsoft.com/en-us/graph/search-concept-overview and
    https://learn.microsoft.com/en-us/graph/search-concept-sort.

    Args:
        client: An authenticated Graph client.
        entity_types: Entity types to search, e.g. `[EntityType.DriveItem]`.
            Only some combinations are allowed by the API.
        options: Query, sort, field, paging, and scope settings.

    Yields:
        Each hit's typed resource, e.g. a `DriveItem` or `Message`. When the
        API leaves the resource ID empty, it is filled from the hit ID.

    Raises:
        ForbiddenGraphError: If the API returns 403 Forbidden.
    """
    page = 0
    total_yielded = 0

    sort_by = options.sort_by

    # https://learn.microsoft.com/en-us/graph/search-concept-sort
    # Note: Sort is not supported for message and event.
    if sort_by is not None and any(
        et in {EntityType.Message, EntityType.Event} for et in entity_types
    ):
        logger.warning(
            "Sorting is not supported for 'message' or 'event'. Disabling sorting for compatibility."
        )
        sort_by = None

    page_size = options.page_size

    while True:
        # Scope query to specific drive if drive_id is provided
        query_string = options.query_string
        if options.drive_id:
            # Add drive ID filter to restrict search to specific drive
            query_string = f"({query_string}) AND DriveId:{options.drive_id}"
            logger.debug(f"Scoping search to drive: {options.drive_id}")

        effective_query = query_string
        if options.site_path:
            effective_query = f'({query_string}) Path:"{options.site_path}"'

        search_request = SearchRequest(
            entity_types=entity_types,
            query=SearchQuery(query_string=effective_query),
            sort_properties=(
                [SortProperty(name=sort_by, is_descending=options.descending)]
                if sort_by
                else None
            ),
            from_=page * page_size,
            size=page_size,
            fields=options.fields,
            region=options.region,
            share_point_one_drive_options=options.share_point_one_drive_options,
        )

        try:
            result = await client.search.query.post(
                body=QueryPostRequestBody(requests=[search_request])
            )

            hits_container = (
                result.value[0].hits_containers[0]
                if result.value[0].hits_containers
                else None
            )
            hits = hits_container.hits or [] if hits_container else []

            if not hits:
                logger.info("No hits found.")
                break

            for hit in hits:
                resource = hit.resource
                # Graph Search does not populate resource.id for Message (and some
                # other entity types); hit.hit_id is the canonical source.
                if resource is not None and not resource.id and hit.hit_id:
                    resource.id = hit.hit_id
                yield resource

            total_yielded += len(hits)
            more = hits_container.more_results_available if hits_container else False
            logger.info(f"Page {page + 1}: {len(hits)} result(s) — {total_yielded} total{' (more available)' if more else ''}")

            # Stop if the server signals no further pages
            if not more:
                break

            # Respect caller-imposed page cap (e.g. Exchange caps from+size≤1000)
            if options.max_pages is not None and page + 1 >= options.max_pages:
                break

            page += 1

        except Exception as exc:
            raise_if_forbidden(exc)
            logger.error(f"Error during search: {exc}")
            break
