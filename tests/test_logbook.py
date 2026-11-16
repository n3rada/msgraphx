"""Tests for msgraphx.utils.logbook module."""

from __future__ import annotations

import json
import os
from collections.abc import Iterator
from pathlib import Path
from unittest import mock

import pytest
from loguru import logger

from msgraphx.utils.logbook import setup_logging


class TestJsonlFileSink:
    """The persistent log file is JSONL: one flat JSON object per line."""

    @pytest.fixture
    def log_file(self, tmp_path: Path) -> Iterator[Path]:
        with mock.patch.dict(
            os.environ, {"MSGRAPHX_LOG_DIR": str(tmp_path)}, clear=False
        ):
            setup_logging("TRACE")
        yield tmp_path / "msgraphx.jsonl"
        # Flushes the enqueued file sink and releases the file handle.
        logger.remove()

    def _read_entries(self, path: Path) -> list[dict]:
        logger.remove()
        lines = path.read_text(encoding="utf-8").splitlines()
        return [json.loads(line) for line in lines]

    def test_every_line_is_json(self, log_file: Path) -> None:
        logger.info("first")
        logger.warning("second")

        messages = [entry["msg"] for entry in self._read_entries(log_file)]

        assert "first" in messages
        assert "second" in messages

    def test_entry_schema(self, log_file: Path) -> None:
        logger.success("schema check")

        entry = next(
            e for e in self._read_entries(log_file) if e["msg"] == "schema check"
        )

        assert entry["level"] == "SUCCESS"
        assert entry["module"] == __name__
        assert entry["function"] == "test_entry_schema"
        assert isinstance(entry["line"], int)
        assert entry["exception"] is None
        assert entry["ts"].endswith("Z")
        assert "extra" not in entry

    def test_multiline_message_stays_on_one_line(self, log_file: Path) -> None:
        logger.info("subject: Q3\nbody: {braces} stay literal")

        entry = next(
            e for e in self._read_entries(log_file) if e["msg"].startswith("subject")
        )

        assert entry["msg"] == "subject: Q3\nbody: {braces} stay literal"

    def test_exception_is_serialized(self, log_file: Path) -> None:
        try:
            raise ValueError("boom")
        except ValueError:
            logger.exception("failed")

        entry = next(e for e in self._read_entries(log_file) if e["msg"] == "failed")

        assert entry["level"] == "ERROR"
        assert "ValueError: boom" in entry["exception"]
        assert "Traceback" in entry["exception"]

    def test_bound_extra_is_included(self, log_file: Path) -> None:
        logger.bind(module_name="outlook", count=42).info("with extra")

        entry = next(
            e for e in self._read_entries(log_file) if e["msg"] == "with extra"
        )

        assert entry["extra"] == {"module_name": "outlook", "count": 42}
