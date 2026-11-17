# tests/test_graph_search.py

from __future__ import annotations

from types import SimpleNamespace
from typing import Any

from msgraph.generated.models.drive_item import DriveItem
from msgraph.generated.models.entity_type import EntityType
from msgraph.generated.models.search_hit import SearchHit
from msgraph.generated.models.search_hits_container import SearchHitsContainer
from msgraph.generated.models.search_response import SearchResponse
from msgraph.generated.search.query.query_post_response import QueryPostResponse

from msgraphx.core.graph_search import SearchOptions, search_entities


def _client(response: Any) -> Any:
    async def post(body: Any) -> Any:
        return response

    return SimpleNamespace(
        search=SimpleNamespace(query=SimpleNamespace(post=post))
    )


async def _collect(response: Any) -> list[Any]:
    return [
        item
        async for item in search_entities(
            _client(response), [EntityType.DriveItem], SearchOptions()
        )
    ]


async def test_empty_response_yields_nothing():
    assert await _collect(None) == []
    assert await _collect(QueryPostResponse(value=[])) == []


async def test_hits_without_resource_are_skipped_and_ids_filled():
    container = SearchHitsContainer(
        hits=[
            SearchHit(hit_id="h1", resource=DriveItem(name="a.txt")),
            SearchHit(hit_id="h2", resource=None),
        ],
        more_results_available=False,
    )
    response = QueryPostResponse(value=[SearchResponse(hits_containers=[container])])

    items = await _collect(response)

    assert [item.name for item in items] == ["a.txt"]
    assert items[0].id == "h1"
