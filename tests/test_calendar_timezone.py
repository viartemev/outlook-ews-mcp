from __future__ import annotations

from types import SimpleNamespace

from exchangelib.ewsdatetime import EWSDateTime, EWSTimeZone

from outlook_mcp.exchange_client import EWSExchangeBackend
from outlook_mcp.models import FindFreeSlotsRequest, ListEventsRequest


class FakeCalendarFolder:
    def __init__(self) -> None:
        self.start = None
        self.end = None

    def view(self, start, end):
        self.start = start
        self.end = end
        return []


def test_list_events_normalizes_pydantic_tzinfo(settings) -> None:
    backend = EWSExchangeBackend(settings)
    folder = FakeCalendarFolder()
    backend._account = SimpleNamespace(
        calendar=folder,
        default_timezone=EWSTimeZone("Europe/Moscow"),
    )

    request = ListEventsRequest.model_validate(
        {
            "start": "2026-04-13T00:00:00+03:00",
            "end": "2026-04-14T00:00:00+03:00",
        }
    )

    assert type(request.start.tzinfo).__module__.startswith("pydantic_core")

    result = backend.list_events(request)

    assert result == []
    assert isinstance(folder.start, EWSDateTime)
    assert isinstance(folder.end, EWSDateTime)
    assert folder.start.tzinfo.key == "Europe/Moscow"
    assert folder.end.tzinfo.key == "Europe/Moscow"


def test_find_free_slots_normalizes_pydantic_tzinfo(settings) -> None:
    backend = EWSExchangeBackend(settings)
    captured: dict = {}

    def get_free_busy_info(**kwargs):
        captured.update(kwargs)
        yield SimpleNamespace(merged="0")

    backend._account = SimpleNamespace(
        default_timezone=EWSTimeZone("Europe/Moscow"),
        protocol=SimpleNamespace(get_free_busy_info=get_free_busy_info),
    )

    request = FindFreeSlotsRequest.model_validate(
        {
            "attendees": ["user@example.com"],
            "duration": 60,
            "start": "2026-04-13T09:00:00+03:00",
            "end": "2026-04-13T10:00:00+03:00",
        }
    )

    assert type(request.start.tzinfo).__module__.startswith("pydantic_core")

    result = backend.find_free_slots(request)

    assert len(result) == 1
    assert isinstance(captured["start"], EWSDateTime)
    assert isinstance(captured["end"], EWSDateTime)
    assert captured["start"].tzinfo.key == "Europe/Moscow"
    assert captured["end"].tzinfo.key == "Europe/Moscow"
    assert result[0].start.tzinfo.key == "Europe/Moscow"
    assert result[0].end.tzinfo.key == "Europe/Moscow"
