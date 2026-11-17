# tests/test_me_pagination.py

from __future__ import annotations

import argparse
from collections.abc import Sequence
from datetime import datetime, timezone
from types import SimpleNamespace
from typing import Any

from msgraph.generated.models.drive_item import DriveItem
from msgraph.generated.models.folder import Folder
from msgraph.generated.models.group import Group
from msgraph.generated.models.person import Person
from msgraph.generated.models.resource_visualization import ResourceVisualization
from msgraph.generated.models.shared_insight import SharedInsight
from msgraph.generated.models.trending import Trending
from msgraph.generated.models.used_insight import UsedInsight
from rich.tree import Tree

from msgraphx.modules.me import drive, groups, people, shared, trending, used


class FakePagedBuilder:
    """Request builder serving fixed pages linked by `@odata.nextLink`."""

    def __init__(self, pages: Sequence[Sequence[Any]]):
        self.pages = pages
        self.fetched: list[int] = []
        self.configs: list[Any] = []

    def _page(self, index: int) -> SimpleNamespace:
        self.fetched.append(index)
        has_next = index + 1 < len(self.pages)
        return SimpleNamespace(
            value=list(self.pages[index]),
            odata_next_link=f"https://graph.test/next/{index + 1}" if has_next else None,
        )

    async def get(self, request_configuration: Any = None) -> SimpleNamespace:
        self.configs.append(request_configuration)
        return self._page(0)

    def with_url(self, url: str) -> SimpleNamespace:
        index = int(url.rsplit("/", 1)[1])

        async def get(request_configuration: Any = None) -> SimpleNamespace:
            return self._page(index)

        return SimpleNamespace(get=get)


def _context(**graph: Any) -> SimpleNamespace:
    return SimpleNamespace(graph_client=SimpleNamespace(**graph))


def _me(**me: Any) -> SimpleNamespace:
    return _context(me=SimpleNamespace(**me))


async def test_people_follows_pages_and_stops_at_top():
    builder = FakePagedBuilder([
        [Person(id="1"), Person(id="2")],
        [Person(id="3"), Person(id="4")],
        [Person(id="5")],
    ])

    rows = await people.fetch(_me(people=builder), top=3)

    assert [r["id"] for r in rows] == ["1", "2", "3"]
    assert builder.fetched == [0, 1]
    assert builder.configs[0].query_parameters.top == 3


async def test_people_page_size_is_capped():
    builder = FakePagedBuilder([[Person(id="1")]])

    await people.fetch(_me(people=builder), top=5000)

    assert builder.configs[0].query_parameters.top == 1000


async def test_groups_collects_every_page_sorted_newest_first():
    old = datetime(2020, 1, 1, tzinfo=timezone.utc)
    new = datetime(2024, 1, 1, tzinfo=timezone.utc)
    builder = FakePagedBuilder([
        [Group(id="a", visibility="Private", created_date_time=old)],
        [Group(id="b", visibility="Public", created_date_time=new)],
    ])
    context = _me(transitive_member_of=SimpleNamespace(graph_group=builder))

    rows = await groups.fetch(context)

    assert [r["id"] for r in rows] == ["b", "a"]
    assert builder.fetched == [0, 1]
    assert builder.configs[0].query_parameters.top == 999


async def test_trending_filters_across_pages_until_top():
    def item(kind: str, title: str) -> Trending:
        return Trending(resource_visualization=ResourceVisualization(type=kind, title=title))

    builder = FakePagedBuilder([
        [item("Word", "w1"), item("Excel", "x1")],
        [item("Word", "w2"), item("Excel", "x2")],
        [item("Excel", "x3")],
    ])
    context = _me(insights=SimpleNamespace(trending=builder))

    rows = await trending.fetch(context, top=2, type_filter="excel")

    assert [(r["rank"], r["title"]) for r in rows] == [(1, "x1"), (2, "x2")]
    assert builder.fetched == [0, 1]


async def test_shared_follows_pages_and_stops_at_top():
    builder = FakePagedBuilder([[SharedInsight(id="1")], [SharedInsight(id="2")], [SharedInsight(id="3")]])
    context = _me(insights=SimpleNamespace(shared=builder))

    rows = await shared.fetch(context, top=2)

    assert len(rows) == 2
    assert builder.fetched == [0, 1]


async def test_used_follows_pages_to_the_end():
    builder = FakePagedBuilder([[UsedInsight(id="1")], [UsedInsight(id="2")]])
    context = _me(insights=SimpleNamespace(used=builder))

    rows = await used.fetch(context, top=25)

    assert len(rows) == 2
    assert builder.fetched == [0, 1]


def _drives(root: DriveItem | None, children: dict[str, FakePagedBuilder]) -> SimpleNamespace:
    async def get_root() -> DriveItem | None:
        return root

    items = SimpleNamespace(
        by_drive_item_id=lambda item_id: SimpleNamespace(children=children[item_id])
    )
    drive_builder = SimpleNamespace(items=items, root=SimpleNamespace(get=get_root))
    return SimpleNamespace(by_drive_id=lambda drive_id: drive_builder)


async def test_run_tree_fails_cleanly_when_root_is_missing():
    context = _context(drives=_drives(root=None, children={}))
    args = argparse.Namespace(path="", depth=3)

    assert await drive._run_tree(context, "drive-1", args) == 1


async def test_build_tree_lists_every_page_and_skips_folders_without_id():
    children = FakePagedBuilder([
        [DriveItem(id="a", name="Alpha", folder=Folder(child_count=1)), DriveItem(id="f1", name="one.txt", size=1)],
        [DriveItem(id=None, name="Ghost", folder=Folder(child_count=0)), DriveItem(id="f2", name="two.txt", size=2)],
    ])
    context = _context(drives=_drives(root=None, children={"root": children}))
    tree = Tree("OneDrive/")

    await drive._build_tree(context, "drive-1", "root", tree, depth=1, max_depth=1)

    labels = [str(child.label) for child in tree.children]
    assert any("Alpha/" in label for label in labels)
    assert any("Ghost/" in label for label in labels)
    assert any("one.txt" in label for label in labels)
    assert any("two.txt" in label for label in labels)
    assert children.fetched == [0, 1]
    alpha = next(child for child in tree.children if "Alpha/" in str(child.label))
    ghost = next(child for child in tree.children if "Ghost/" in str(child.label))
    assert [str(c.label) for c in alpha.children] == ["[dim]…[/dim]"]
    assert ghost.children == []
