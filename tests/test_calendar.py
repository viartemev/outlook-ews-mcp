from __future__ import annotations

from datetime import UTC, datetime

import pytest

from outlook_mcp.errors import APIError
from outlook_mcp.tools.calendar import create_event, find_free_slots, get_event, list_events


def test_list_events(client) -> None:
    result = list_events(
        client,
        {
            "start": datetime(2026, 4, 8, 9, 0, tzinfo=UTC).isoformat(),
            "end": datetime(2026, 4, 8, 10, 0, tzinfo=UTC).isoformat(),
        },
    )
    assert result[0]["subject"] == "Planning"


def test_get_event(client) -> None:
    result = get_event(client, {"id": "event-1"})
    assert result["id"] == "event-1"


def test_create_event_validates_range(client) -> None:
    with pytest.raises(APIError) as excinfo:
        create_event(
            client,
            {
                "subject": "Broken",
                "start": datetime(2026, 4, 8, 10, 0, tzinfo=UTC).isoformat(),
                "end": datetime(2026, 4, 8, 9, 0, tzinfo=UTC).isoformat(),
            },
        )

    assert excinfo.value.code == "validation_error"


def test_find_free_slots(client) -> None:
    result = find_free_slots(
        client,
        {
            "attendees": ["user@example.com"],
            "duration": 60,
            "start": datetime(2026, 4, 8, 9, 0, tzinfo=UTC).isoformat(),
            "end": datetime(2026, 4, 8, 18, 0, tzinfo=UTC).isoformat(),
        },
    )
    assert result[0]["all_available"] is True
