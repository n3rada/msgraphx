# tests/test_sharepoint_download.py

from __future__ import annotations

from pathlib import Path
from unittest.mock import AsyncMock, MagicMock

import pytest
from msgraph.generated.models.drive_item import DriveItem
from msgraph.generated.models.file import File
from msgraph.generated.models.folder import Folder
from msgraph.generated.models.o_data_errors.main_error import MainError
from msgraph.generated.models.o_data_errors.o_data_error import ODataError

from msgraphx.modules.sharepoint import download
from msgraphx.utils.errors import AuthenticationError


def _auth_error() -> ODataError:
    return ODataError(
        error=MainError(code="InvalidAuthenticationToken", message="Token expired."),
        response_status_code=401,
    )


def _client_with_content(content: AsyncMock) -> MagicMock:
    client = MagicMock()
    item_builder = client.drives.by_drive_id.return_value.items.by_drive_item_id.return_value
    item_builder.content.get = content
    return client


async def test_nameless_item_is_skipped(tmp_path: Path):
    client = MagicMock()
    item = DriveItem(id="i-1", name=None, file=File())

    assert await download.download_drive_item(client, "d-1", item, tmp_path) == 0
    client.drives.by_drive_id.assert_not_called()


async def test_file_is_written(tmp_path: Path):
    client = _client_with_content(AsyncMock(return_value=b"data"))
    item = DriveItem(id="i-1", name="a.txt", size=4, file=File())

    assert await download.download_drive_item(client, "d-1", item, tmp_path) == 1
    assert (tmp_path / "a.txt").read_bytes() == b"data"


async def test_auth_error_on_file_download_propagates(tmp_path: Path):
    client = _client_with_content(AsyncMock(side_effect=_auth_error()))
    item = DriveItem(id="i-1", name="a.txt", size=4, file=File())

    with pytest.raises(AuthenticationError):
        await download.download_drive_item(client, "d-1", item, tmp_path)


async def test_other_errors_are_logged_not_raised(tmp_path: Path):
    client = _client_with_content(AsyncMock(side_effect=RuntimeError("boom")))
    item = DriveItem(id="i-1", name="a.txt", size=4, file=File())

    assert await download.download_drive_item(client, "d-1", item, tmp_path) == 0


async def test_auth_error_in_folder_child_stops_dump(tmp_path: Path, monkeypatch):
    children = [
        DriveItem(id="c-1", name="ok.txt", size=2, file=File()),
        DriveItem(id="c-2", name="bad.txt", size=2, file=File()),
    ]
    monkeypatch.setattr(download.pagination, "collect_all", AsyncMock(return_value=children))

    client = _client_with_content(AsyncMock(side_effect=[b"ok", _auth_error()]))
    folder = DriveItem(id="f-1", name="Docs", folder=Folder())

    with pytest.raises(AuthenticationError):
        await download.download_drive_item(client, "d-1", folder, tmp_path)


def test_sum_results_counts_and_reraises_auth():
    auth = AuthenticationError("expired")
    assert download._sum_results([1, 2, RuntimeError("x")], "here") == 3
    with pytest.raises(AuthenticationError):
        download._sum_results([1, auth, RuntimeError("x")], "here")


async def test_download_drive_handles_missing_drive(tmp_path: Path):
    client = MagicMock()
    client.drives.by_drive_id.return_value.get = AsyncMock(return_value=None)

    assert await download.download_drive(client, "d-1", tmp_path) == 0


async def test_download_drive_auth_error_propagates(tmp_path: Path):
    client = MagicMock()
    client.drives.by_drive_id.return_value.get = AsyncMock(side_effect=_auth_error())

    with pytest.raises(ODataError):
        await download.download_drive(client, "d-1", tmp_path)
