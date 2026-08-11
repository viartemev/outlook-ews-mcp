from __future__ import annotations

from datetime import UTC, datetime
from types import SimpleNamespace

import pytest
from pydantic import ValidationError

from outlook_mcp.exchange_client import EWSExchangeBackend
from outlook_mcp.models import (
    CalendarEvent,
    CreateEventRequest,
    EmailAddress,
    ListEventsRequest,
)

MOMENT = datetime(2026, 4, 8, 10, 0, tzinfo=UTC)


def _item(start: datetime, end: datetime, item_id: str = "event-1") -> SimpleNamespace:
    return SimpleNamespace(
        id=item_id,
        subject="Согласование",
        start=start,
        end=end,
        organizer=SimpleNamespace(email_address="organizer@example.com", name="Organizer"),
    )


class _EventQuery(list):
    """folder.view() result: list-like, accepting the .only() projection."""

    def only(self, *fields):
        return self


class FakeCalendarFolder:
    def __init__(self, items: list) -> None:
        self.items = items

    def view(self, start, end):
        return _EventQuery(self.items)


def _backend_over(settings, items: list) -> EWSExchangeBackend:
    backend = EWSExchangeBackend(settings)
    backend._account = SimpleNamespace(calendar=FakeCalendarFolder(items))
    return backend


def _window() -> ListEventsRequest:
    return ListEventsRequest(
        start=datetime(2026, 4, 8, 0, 0, tzinfo=UTC),
        end=datetime(2026, 4, 9, 0, 0, tzinfo=UTC),
    )


def test_list_events_keeps_zero_duration_event(settings, monkeypatch) -> None:
    """Exchange stores these; refusing to parse one hid the whole day.

    Placeholders and items created by external systems routinely arrive with
    end == start. list_events converts the whole page in one comprehension, so
    a single such appointment used to fail every event around it.
    """
    backend = _backend_over(settings, [_item(MOMENT, MOMENT)])
    monkeypatch.setattr(backend, "_to_ews_datetime", lambda value: value)

    events = backend.list_events(_window())

    assert [event.id for event in events] == ["event-1"]
    assert events[0].start == events[0].end


def test_list_events_keeps_neighbours_of_a_broken_event(settings, monkeypatch) -> None:
    backend = _backend_over(
        settings,
        [
            _item(MOMENT, MOMENT, "zero"),
            _item(MOMENT, datetime(2026, 4, 8, 11, 0, tzinfo=UTC), "normal"),
        ],
    )
    monkeypatch.setattr(backend, "_to_ews_datetime", lambda value: value)

    assert [event.id for event in backend.list_events(_window())] == ["zero", "normal"]


def test_calendar_event_accepts_inverted_range() -> None:
    """Reversed timestamps are corrupt data, not a reason to hide the calendar.

    The model describes what the server holds. Reporting it lets the user see
    and fix the item; rejecting it only removes the evidence.
    """
    event = CalendarEvent(
        id="event-1",
        subject="Согласование",
        start=MOMENT,
        end=datetime(2026, 4, 8, 9, 0, tzinfo=UTC),
        organizer=EmailAddress(email="organizer@example.com"),
    )
    assert event.end < event.start


def test_creating_an_event_stays_strict() -> None:
    """Leniency must not leak into what we hand to Exchange.

    A zero-length or reversed request is a user mistake, and the moment to say
    so is before the invitation goes out.
    """
    with pytest.raises(ValidationError):
        CreateEventRequest(subject="s", start=MOMENT, end=MOMENT)

    with pytest.raises(ValidationError):
        CreateEventRequest(subject="s", start=MOMENT, end=datetime(2026, 4, 8, 9, 0, tzinfo=UTC))


def test_listing_a_reversed_window_stays_strict() -> None:
    with pytest.raises(ValidationError):
        ListEventsRequest(start=MOMENT, end=MOMENT)


def test_free_slots_ignore_a_zero_duration_event(settings) -> None:
    """A zero-length appointment blocks nothing, and must not split the day."""
    backend = EWSExchangeBackend(settings)
    day_start = datetime(2026, 4, 8, 9, 0, tzinfo=UTC)
    day_end = datetime(2026, 4, 8, 18, 0, tzinfo=UTC)
    event = CalendarEvent(
        id="event-1",
        subject="Согласование",
        start=MOMENT,
        end=MOMENT,
        organizer=EmailAddress(email="organizer@example.com"),
    )

    slots = backend._compute_free_slots(day_start, day_end, [event])

    assert [(slot.start, slot.end) for slot in slots] == [(day_start, day_end)]
