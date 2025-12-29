# msgraphx/utils/pagination.py

# Built-in imports
from typing import TypeVar, AsyncIterator, Any, Protocol
from loguru import logger
from .errors import handle_graph_auth_errors

T = TypeVar("T")


class RequestBuilder(Protocol):
    """Protocol for Graph SDK request builders."""

    async def get(self, request_configuration: Any = None) -> Any:
        """Get request that returns paginated results."""
        ...

    def with_url(self, url: str) -> "RequestBuilder":
        """Create builder with next page URL."""
        ...


class GraphPaginator:
    """
    Modern async iterator for Microsoft Graph API pagination.

    Implements proper async iterator protocol for clean, Pythonic pagination.
    Works seamlessly with Python's `async for` syntax.

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
        request_configuration: Any = None,
        max_pages: int | None = None,
    ):
        """
        Initialize paginator.

        Args:
            request_builder: Graph SDK request builder (e.g., client.users, client.groups)
            request_configuration: Optional request configuration with query parameters
            max_pages: Optional limit on number of pages to fetch
        """
        self._builder = request_builder
        self._config = request_configuration
        self._max_pages = max_pages
        self._current_result = None
        self._current_index = 0
        self._page_count = 0
        self._started = False

    def __aiter__(self) -> AsyncIterator[T]:
        """Return self as async iterator."""
        return self

    @handle_graph_auth_errors
    async def __anext__(self) -> T:
        """Get next item from paginated results."""
        # First iteration - get initial page
        if not self._started:
            self._current_result = await self._builder.get(
                request_configuration=self._config
            )
            self._started = True

            if not self._current_result or not self._current_result.value:
                raise StopAsyncIteration

            self._page_count = 1
            logger.debug(
                f"📄 Page {self._page_count}: {len(self._current_result.value)} items"
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
            logger.warning(f"⚠️ Reached maximum page limit: {self._max_pages}")
            raise StopAsyncIteration

        # Fetch next page
        self._current_result = await self._builder.with_url(
            self._current_result.odata_next_link
        ).get()

        if not self._current_result or not self._current_result.value:
            raise StopAsyncIteration

        self._page_count += 1
        logger.debug(
            f"📄 Page {self._page_count}: {len(self._current_result.value)} items"
        )
        self._current_index = 0

        # Return first item from new page
        item = self._current_result.value[self._current_index]
        self._current_index += 1
        return item

    async def collect(self) -> list[T]:
        """Collect all results into a list."""
        results = []
        async for item in self:
            results.append(item)
        return results

    async def filter(self, predicate) -> list[T]:
        """Collect only items matching predicate."""
        results = []
        async for item in self:
            if predicate(item):
                results.append(item)
        return results

    async def count(self) -> int:
        """Count total items (fetches all pages)."""
        count = 0
        async for _ in self:
            count += 1
        return count


# Convenience function for quick collection
async def collect_all(
    request_builder: RequestBuilder,
    request_configuration: Any = None,
    max_pages: int | None = None,
) -> list[T]:
    """
    Convenience function to collect all paginated results.

    Args:
        request_builder: Graph SDK request builder
        request_configuration: Optional request configuration
        max_pages: Optional page limit

    Returns:
        List of all items

    Example:
        all_groups = await collect_all(
            client.groups,
            GroupsRequestBuilder.GroupsRequestBuilderGetRequestConfiguration(
                query_parameters=query_params
            )
        )
    """
    return await GraphPaginator(
        request_builder, request_configuration, max_pages
    ).collect()
