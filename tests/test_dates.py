"""Tests for msgraphx.utils.dates module."""

from __future__ import annotations

import re
from datetime import datetime, timezone, timedelta

import pytest

from msgraphx.utils.dates import parse_date_string

ISO_PATTERN = re.compile(r"^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}Z$")


class TestRelativeDurations:
    """Test relative duration parsing (e.g. '5h', '2d', '1w')."""

    @pytest.mark.parametrize(
        "input_str,unit_seconds",
        [
            ("30s", 30),
            ("5m", 300),
            ("2h", 7200),
            ("1d", 86400),
            ("1w", 604800),
            ("1y", 365 * 86400),
        ],
    )
    def test_relative_units(self, input_str, unit_seconds):
        result = parse_date_string(input_str)
        assert ISO_PATTERN.match(result)

        parsed = datetime.strptime(result, "%Y-%m-%dT%H:%M:%SZ").replace(
            tzinfo=timezone.utc
        )
        now = datetime.now(timezone.utc)
        # Allow 2 second tolerance for test execution time
        expected = now - timedelta(seconds=unit_seconds)
        assert abs((parsed - expected).total_seconds()) < 2

    def test_large_day_value(self):
        result = parse_date_string("30d")
        assert ISO_PATTERN.match(result)

    def test_case_insensitive(self):
        lower = parse_date_string("5d")
        upper = parse_date_string("5D")
        # Both should produce the same result (within 1 second)
        t_lower = datetime.strptime(lower, "%Y-%m-%dT%H:%M:%SZ")
        t_upper = datetime.strptime(upper, "%Y-%m-%dT%H:%M:%SZ")
        assert abs((t_lower - t_upper).total_seconds()) < 1


class TestAbsoluteDates:
    """Test absolute date parsing (YYYY-MM-DD)."""

    def test_valid_date(self):
        result = parse_date_string("2024-06-15")
        assert result == "2024-06-15T00:00:00Z"

    def test_start_of_year(self):
        result = parse_date_string("2025-01-01")
        assert result == "2025-01-01T00:00:00Z"

    def test_end_of_year(self):
        result = parse_date_string("2024-12-31")
        assert result == "2024-12-31T00:00:00Z"


class TestInvalidInputs:
    """Test that invalid inputs raise ValueError."""

    @pytest.mark.parametrize(
        "invalid",
        [
            "abc",
            "5x",  # invalid unit
            "2024/06/15",  # wrong separator
            "15-06-2024",  # wrong order
            "",
            "5",  # missing unit
            "d5",  # reversed
        ],
    )
    def test_invalid_raises(self, invalid):
        with pytest.raises(ValueError):
            parse_date_string(invalid)
