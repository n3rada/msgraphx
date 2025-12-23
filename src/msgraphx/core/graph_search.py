# Built-in imports
from collections.abc import AsyncGenerator
from dataclasses import dataclass
from typing import Optional

# External library imports
from loguru import logger
from msgraph import GraphServiceClient
from msgraph.generated.models.search_request import SearchRequest
from msgraph.generated.models.search_query import SearchQuery
from msgraph.generated.models.sort_property import SortProperty
from msgraph.generated.search.query.query_post_request_body import QueryPostRequestBody
from msgraph.generated.models.entity_type import EntityType


@dataclass
class SearchOptions:
    query_string: str = "*"
    sort_by: Optional[str] = "createdDateTime"
    descending: bool = True
    fields: Optional[list[str]] = None
    page_size: int = 500
    region: Optional[str] = None  # Required for application permissions
    drive_id: Optional[str] = None  # Scope search to specific drive


async def search_entities(
    client: GraphServiceClient,
    entity_types: list[EntityType],
    options: SearchOptions = SearchOptions(),
) -> AsyncGenerator[EntityType, None]:
    """
    Perform a paginated Microsoft Graph search across one or more entity types (e.g., DriveItem, Message, Site).

    Supports optional sorting (where allowed) and field filtering. This function handles pagination
    transparently and yields results across all pages.

    Args:
        client (GraphServiceClient): An authenticated Microsoft Graph client.
        entity_types (list[EntityType]): The types of entities to search (e.g., [EntityType.DriveItem]).
                                         Refer to Microsoft documentation for supported combinations.
        options (SearchOptions, optional): Search configuration including query string, sorting, fields, and page size.

    Yields:
        EntityType: Each yielded object is the typed resource (e.g., DriveItem, Message) returned by the Graph Search API.

    Warnings:
        If sorting is requested on unsupported entity types (e.g., Message or Event), sorting will be disabled
        and a warning will be logged.

    References:
        https://learn.microsoft.com/en-us/graph/search-concept-overview
        https://learn.microsoft.com/en-us/graph/search-concept-sort

    This is a POST request to the `/search/query` endpoint, which allows for complex queries such as:

    ```
    {
        "requests": [
            {
            "entityTypes": ["driveItem"],
            "query": {
                "queryString": "filetype:pdf"
            },
            "sortProperties": [
                {
                "name": "createdDateTime",
                "isDescending": true
                }
            ],
            "from": 0,
            "size": 20
            }
        ]
    }
    ```

    """
    page = 0

    sort_by = options.sort_by

    # https://learn.microsoft.com/en-us/graph/search-concept-sort
    # Note: Sort is not supported for message and event.
    if sort_by is not None and any(
        et in {EntityType.Message, EntityType.Event} for et in entity_types
    ):
        logger.warning(
            "⚠️ Sorting is not supported for 'message' or 'event'. Disabling sorting for compatibility."
        )
        sort_by = None

    page_size = options.page_size

    while True:
        # Scope query to specific drive if drive_id is provided
        query_string = options.query_string
        if options.drive_id:
            # Add drive ID filter to restrict search to specific drive
            query_string = f"({query_string}) AND DriveId:{options.drive_id}"
            logger.debug(f"🔒 Scoping search to drive: {options.drive_id}")

        search_request = SearchRequest(
            entity_types=entity_types,
            query=SearchQuery(query_string=query_string),
            sort_properties=(
                [SortProperty(name=sort_by, is_descending=options.descending)]
                if sort_by
                else None
            ),
            from_=page * page_size,
            size=page_size,
            fields=options.fields,
            region=options.region,
        )

        try:
            result = await client.search.query.post(
                body=QueryPostRequestBody(requests=[search_request])
            )

            hits = (
                (result.value[0].hits_containers[0].hits or [])
                if result.value[0].hits_containers
                else []
            )

            if not hits:
                logger.info("📭 No hits found.")
                break

            for hit in hits:
                yield hit.resource

            page += 1

        except Exception as exc:
            logger.error(f"❌ Error during search: {exc}")
            break
