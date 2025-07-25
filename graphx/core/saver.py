# Built-in imports
import re
import os
from pathlib import Path
from datetime import datetime, timezone
from urllib.parse import unquote

# External library imports
from loguru import logger

# Local library imports
from msgraph import GraphServiceClient
from msgraph.generated.models.entity_type import EntityType


def sanitize_filename(name: str) -> str:
    return re.sub(r'[<>:"/\\|?*\x00-\x1F]', "_", name)


class Saver:

    def __init__(self, graph_client: GraphServiceClient, drop_folder: str = None):

        self._graph_client = graph_client

        if drop_folder is None:
            utc_day = datetime.now(tz=timezone.utc).strftime("%Y-%m-%d")
            drop_folder = Path.cwd() / f"graphx_{utc_day}"
        else:
            drop_folder = Path(drop_folder)

        self.__drop_folder = drop_folder

        try:
            self.__drop_folder.mkdir(parents=True, exist_ok=True)
        except Exception as exc:
            raise PermissionError(f"Failed to create drop folder: {exc}") from exc

        # Check write permission by attempting to write a temp file
        try:
            test_file = self.__drop_folder / ".write_test"
            test_file.write_bytes("w")
            test_file.unlink()
        except Exception as exc:
            raise PermissionError(f"Cannot write to {self.__drop_folder}") from exc

        logger.success(
            f"💾 Initialized saver. Files will be stored in: {self.__drop_folder}"
        )

    def save(self, content: bytes | str, path: str | Path) -> None:
        if not path:
            raise ValueError("Path must not be empty")

        try:
            path = Path(path)
            if not str(path).startswith(str(self.__drop_folder)):
                full_path = self.__drop_folder / path
            else:
                full_path = path

            if full_path.suffix == "":
                full_path.mkdir(parents=True, exist_ok=True)
                return

            full_path.parent.mkdir(parents=True, exist_ok=True)

            if isinstance(content, str):
                full_path.write_text(content, encoding="utf-8")
            elif isinstance(content, bytes):
                full_path.write_bytes(content)
            else:
                raise TypeError("Content must be str or bytes")

        except (IOError, TypeError, ValueError) as e:
            logger.error(f"❌ Failed to save content to {full_path}: {e}")
            raise

    def save_drive_item(
        self, item: EntityType.DriveItem, rebuild_sharepoint_path: bool = False
    ) -> None:
        item_name = item.name
        item_url = item.web_url

        drive_id = item.parent_reference.drive_id

        if rebuild_sharepoint_path:

            # Get the portion after /sites/
            # ClientAnalyticsPlatform/Shared Documents/Project Okta Channel/Okta Client Migration Resources/MFA ADMIN Email.docx.url
            site_relative = unquote(item_url.split("/sites/", 1)[-1])

            # Directory path
            parts = site_relative.strip("/").split("/")

            folder_parts = [sanitize_filename(p) for p in parts[:-1]]  # Skip item name
            dir_path = os.path.join(self.__drop_folder, "sites", *folder_parts)
        else:
            dir_path = str(self.__drop_folder)

        save_path = os.path.join(dir_path, item_name)
