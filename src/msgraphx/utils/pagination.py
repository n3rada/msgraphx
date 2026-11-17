# msgraphx/utils/pagination.py

# Built-in imports
from __future__ import annotations

from collections.abc import AsyncIterator, Callable
from typing import Any, Protocol

from kiota_abstractions.base_request_configuration import RequestConfiguration
from loguru import logger
from msgraph.generated.models.o_data_errors.o_data_error import ODataError

from .errors import handle_graph_errors

_TRANSIENT_STATUS_CODES: frozenset[int] = frozenset({429, 500, 502, 503, 504})


class RequestBuilder(Protocol):
    """A Graph SDK request builder for a paginated collection.

    Generated builders such as `client.users` or `client.groups` match this
    protocol structurally. Collection responses differ per builder and share
    no base class, so items and responses are typed `Any`.
    """

    async def get(
        self, request_configuration: RequestConfiguration[Any] | None = None
    ) -> Any:
        """Fetch one page; the response has `value` and `odata_next_link`."""
        ...

    def with_url(self, url: str) -> RequestBuilder:
        """Return a builder that targets `url`, e.g. an `@odata.nextLink`."""
        ...


class GraphPaginator:
    """An async iterator over every item of a paginated Graph collection.

    Follows `@odata.nextLink` until the last page, an empty page, or
    `max_pages`. A 429 or 5xx response logs an error and ends iteration
    early instead of raising. Authentication failures are raised as
    `AuthenticationError`.

    Example:
        # Simple iteration
        paginator = GraphPaginator(client.users, request_config)
        async for user in paginator:
            print(user.display_name)

        # Collect all at once
        all_users = await GraphPaginator(client.users, request_config).collect()

        # With filter
        async for group in GraphPaginator(client.groups, query_config):
            if group.security_enabled:
                print(group.display_name)
    """

    def __init__(
        self,
        request_builder: RequestBuilder,
        request_configuration: RequestConfiguration[Any] | None = None,
        max_pages: int | None = None,
    ):
        """Prepare iteration; no request is sent until the first item.

        Args:
            request_builder: SDK request builder for the collection, e.g.
                `client.users`.
            request_configuration: Query parameters and headers for the
                first request. Later pages follow `@odata.nextLink`.
            max_pages: Stop after this many pages. `None` fetches all pages.
        """
        self._builder = request_builder
        self._config = request_configuration
        self._max_pages = max_pages
        # The SDK collection response for the current page.
        self._current_result: Any = None
        self._current_index = 0
        self._page_count = 0
        self._started = False

    def __aiter__(self) -> AsyncIterator[Any]:
        """Return this paginator; iterating it consumes it."""
        return self

    @handle_graph_errors
    async def __anext__(self) -> Any:
        """Return the next item, fetching the next page when needed.

        Raises:
            StopAsyncIteration: When the collection is exhausted, the page
                limit is reached, or the API returned 429 or 5xx.
            AuthenticationError: If the token is invalid or expired.
        """
        # First iteration - get initial page
        if not self._started:
            try:
                self._current_result = await self._builder.get(
                    request_configuration=self._config
                )
            except ODataError as exc:
                status = getattr(exc, "response_status_code", None)
                if status in _TRANSIENT_STATUS_CODES:
                    logger.error(
                        f"HTTP {status} from Graph API, aborting pagination"
                    )
                    raise StopAsyncIteration from exc
                raise
            self._started = True

            if not self._current_result or not self._current_result.value:
                raise StopAsyncIteration

            self._page_count = 1
            logger.debug(
                f"Page {self._page_count}: {len(self._current_result.value)} items"
            )
            self._current_index = 0

        # Return next item from current page
        if self._current_index < len(self._current_result.value):
            item = self._current_result.value[self._current_index]
            self._current_index += 1
            return item

        # Current page exhausted, check for next page
        if self._current_result.odata_next_link is None:
            raise StopAsyncIteration

        # Check page limit
        if self._max_pages and self._page_count >= self._max_pages:
            logger.warning(f"Reached maximum page limit: {self._max_pages}")
            raise StopAsyncIteration

        # Fetch next page
        next_link = self._current_result.odata_next_link
        try:
            self._current_result = await self._builder.with_url(next_link).get()
        except ODataError as exc:
            status = getattr(exc, "response_status_code", None)
            if status in _TRANSIENT_STATUS_CODES:
                logger.error(f"HTTP {status} from Graph API, aborting pagination")
                raise StopAsyncIteration from exc
            raise

        if not self._current_result or not self._current_result.value:
            raise StopAsyncIteration

        self._page_count += 1
        logger.debug(
            f"Page {self._page_count}: {len(self._current_result.value)} items"
        )
        self._current_index = 0

        # Return first item from new page
        item = self._current_result.value[self._current_index]
        self._current_index += 1
        return item

    async def collect(self) -> list[Any]:
        """Return every remaining item, fetching all remaining pages."""
        results: list[Any] = []
        async for item in self:
            results.append(item)
        return results

    async def filter(self, predicate: Callable[[Any], bool]) -> list[Any]:
        """Return the remaining items that satisfy a predicate.

        Args:
            predicate: Called with each item; truthy keeps the item.

        Returns:
            The kept items, in server order.
        """
        results: list[Any] = []
        async for item in self:
            if predicate(item):
                results.append(item)
        return results

    async def count(self) -> int:
        """Return the number of remaining items, fetching all pages."""
        count = 0
        async for _ in self:
            count += 1
        return count


# Convenience function for quick collection
async def collect_all(
    request_builder: RequestBuilder,
    request_configuration: RequestConfiguration[Any] | None = None,
    max_pages: int | None = None,
) -> list[Any]:
    """Collect every item from a paginated Graph collection.

    Args:
        request_builder: SDK request builder for the collection, e.g.
            `client.groups`.
        request_configuration: Query parameters and headers for the first
            request. Later pages follow `@odata.nextLink`.
        max_pages: Stop after this many pages. `None` fetches all pages.

    Returns:
        All items across the fetched pages, in server order.

    Raises:
        AuthenticationError: If the token is invalid or expired.

    Example:
        all_groups = await collect_all(
            client.groups,
            RequestConfiguration(query_parameters=query_params)
        )
    """
    return await GraphPaginator(
        request_builder, request_configuration, max_pages
    ).collect()
