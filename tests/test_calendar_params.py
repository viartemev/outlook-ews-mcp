from __future__ import annotations

from datetime import UTC, datetime, date
from types import SimpleNamespace

import pytest
from exchangelib.errors import UnauthorizedError

import outlook_mcp.exchange_client.calendar as calendar_module
from outlook_mcp.errors import APIError, AuthFailedError
from outlook_mcp.exchange_client import EWSExchangeBackend
from outlook_mcp.models import (
    CreateEventRequest,
    DeleteEventRequest,
    FindFreeSlotsRequest,
    GetEventRequest,
    ListEventsRequest,
    RecurrencePattern,
    WorkHours,
)


class FakeCalendarItem:
    saved: list["FakeCalendarItem"] = []

    def __init__(self, **kwargs) -> None:
        self.kwargs = kwargs
        self.id = "event-new"

    def save(self, **kwargs) -> None:
        FakeCalendarItem.saved.append(self)


def test_create_event_passes_recurrence(settings, monkeypatch) -> None:
    monkeypatch.setattr(calendar_module, "CalendarItem", FakeCalendarItem)
    backend = EWSExchangeBackend(settings)
    backend._account = SimpleNamespace(calendar=object(), default_timezone=UTC)

    request = CreateEventRequest.model_validate(
        {
            "subject": "Sprint planning",
            "start": "2026-04-13T09:00:00+00:00",
            "end": "2026-04-13T10:00:00+00:00",
            "recurrence": {
                "type": "weekly",
                "interval": 2,
                "days_of_week": ["monday", "wednesday"],
            },
        }
    )

    backend.create_event(request)

    saved = FakeCalendarItem.saved[-1]
    recurrence = saved.kwargs["recurrence"]
    assert recurrence.pattern.interval == 2
    assert list(recurrence.pattern.weekdays) == ["Monday", "Wednesday"]


def test_create_event_without_recurrence_leaves_it_unset(settings, monkeypatch) -> None:
    monkeypatch.setattr(calendar_module, "CalendarItem", FakeCalendarItem)
    backend = EWSExchangeBackend(settings)
    backend._account = SimpleNamespace(calendar=object(), default_timezone=UTC)

    request = CreateEventRequest.model_validate(
        {
            "subject": "One-off",
            "start": "2026-04-13T09:00:00+00:00",
            "end": "2026-04-13T10:00:00+00:00",
        }
    )

    backend.create_event(request)

    saved = FakeCalendarItem.saved[-1]
    assert saved.kwargs["recurrence"] is None


class FakeCalendarFolder:
    def __init__(self, items: list) -> None:
        self._items = items

    def view(self, start, end):
        return self._items


def _fake_event(
    item_id: str, is_recurring: bool, parent_folder_id: str | None = None
) -> SimpleNamespace:
    return SimpleNamespace(
        id=item_id,
        subject="Standup",
        start=datetime(2026, 4, 13, 9, 0, tzinfo=UTC),
        end=datetime(2026, 4, 13, 9, 30, tzinfo=UTC),
        is_recurring=is_recurring,
        parent_folder_id=SimpleNamespace(id=parent_folder_id) if parent_folder_id else None,
    )


def test_list_events_excludes_recurring_when_requested(settings) -> None:
    backend = EWSExchangeBackend(settings)
    folder = FakeCalendarFolder([_fake_event("single", False), _fake_event("recurring", True)])
    backend._account = SimpleNamespace(calendar=folder, default_timezone=UTC)

    request = ListEventsRequest.model_validate(
        {
            "start": "2026-04-13T00:00:00+00:00",
            "end": "2026-04-14T00:00:00+00:00",
            "include_recurring": False,
        }
    )

    result = backend.list_events(request)

    assert [event.id for event in result] == ["single"]


def test_list_events_includes_recurring_by_default(settings) -> None:
    backend = EWSExchangeBackend(settings)
    folder = FakeCalendarFolder([_fake_event("single", False), _fake_event("recurring", True)])
    backend._account = SimpleNamespace(calendar=folder, default_timezone=UTC)

    request = ListEventsRequest.model_validate(
        {
            "start": "2026-04-13T00:00:00+00:00",
            "end": "2026-04-14T00:00:00+00:00",
        }
    )

    result = backend.list_events(request)

    assert {event.id for event in result} == {"single", "recurring"}


class FakeDeletableItem:
    def __init__(self) -> None:
        self.deleted_with: dict | None = None
        self.cancelled_with: dict | None = None
        self.parent_folder_id = SimpleNamespace(id="cal-1")

    def delete(self, **kwargs) -> None:
        self.deleted_with = kwargs

    def cancel(self, **kwargs) -> None:
        self.cancelled_with = kwargs


def test_delete_event_sends_cancel_message_when_provided(settings) -> None:
    backend = EWSExchangeBackend(settings)
    item = FakeDeletableItem()
    backend._account = SimpleNamespace(
        calendar=SimpleNamespace(id="cal-1"),
        fetch=lambda **kwargs: iter([item]),
    )

    request = DeleteEventRequest.model_validate(
        {"id": "event-1", "notify_attendees": True, "cancel_message": "Meeting cancelled, sorry!"}
    )

    backend.delete_event(request)

    assert item.cancelled_with == {"body": "Meeting cancelled, sorry!"}
    assert item.deleted_with is None


def test_delete_event_without_message_still_deletes(settings) -> None:
    backend = EWSExchangeBackend(settings)
    item = FakeDeletableItem()
    backend._account = SimpleNamespace(
        calendar=SimpleNamespace(id="cal-1"),
        fetch=lambda **kwargs: iter([item]),
    )

    request = DeleteEventRequest.model_validate({"id": "event-1", "notify_attendees": True})

    backend.delete_event(request)

    assert item.deleted_with == {"send_meeting_cancellations": "SendToAllAndSaveCopy"}
    assert item.cancelled_with is None


def test_create_event_all_day_same_day_becomes_one_ews_day(settings, monkeypatch) -> None:
    """A same-day start/end marked is_all_day must become a single EWS day
    (exclusive end date), not a zero-duration item."""
    monkeypatch.setattr(calendar_module, "CalendarItem", FakeCalendarItem)
    backend = EWSExchangeBackend(settings)
    backend._account = SimpleNamespace(calendar=object(), default_timezone=UTC)

    request = CreateEventRequest.model_validate(
        {
            "subject": "Day off",
            "start": "2026-08-03T09:00:00+00:00",
            "end": "2026-08-03T17:00:00+00:00",
            "is_all_day": True,
        }
    )

    result = backend.create_event(request)

    assert result.start == datetime(2026, 8, 3, 0, 0, tzinfo=UTC)
    assert result.end == datetime(2026, 8, 4, 0, 0, tzinfo=UTC)


def test_create_event_all_day_does_not_add_an_extra_day(settings, monkeypatch) -> None:
    """Regression: '3-4 August' used to become 03.08 00:00 - 04.08 23:59:59 --
    almost two days -- instead of the single EWS day the exclusive-end
    convention implies."""
    monkeypatch.setattr(calendar_module, "CalendarItem", FakeCalendarItem)
    backend = EWSExchangeBackend(settings)
    backend._account = SimpleNamespace(calendar=object(), default_timezone=UTC)

    request = CreateEventRequest.model_validate(
        {
            "subject": "Day off",
            "start": "2026-08-03T00:00:00+00:00",
            "end": "2026-08-04T00:00:00+00:00",
            "is_all_day": True,
        }
    )

    result = backend.create_event(request)

    assert result.start == datetime(2026, 8, 3, 0, 0, tzinfo=UTC)
    assert result.end == datetime(2026, 8, 4, 0, 0, tzinfo=UTC)


def test_get_event_uses_alternate_calendar_when_calendar_id_given(settings, monkeypatch) -> None:
    default_calendar = SimpleNamespace(id="cal-default")
    alt_calendar = SimpleNamespace(id="cal-2")
    fetch_calls: list = []
    backend = EWSExchangeBackend(settings)

    def fetch(ids, folder=None):
        fetch_calls.append(folder)
        # A real fetch resolves the item wherever it actually lives -- match
        # parent_folder_id to whichever folder this call targeted so the
        # scoping check in _fetch_item passes the same way a real EWS lookup
        # of an item that really is in that folder would.
        return iter([_fake_event("event-1", False, parent_folder_id=folder.id)])

    backend._account = SimpleNamespace(calendar=default_calendar, fetch=fetch)
    monkeypatch.setattr(
        backend,
        "_resolve_folder",
        lambda value: alt_calendar if value == "cal-2" else default_calendar,
    )

    backend.get_event(GetEventRequest(id="event-1", calendar_id="cal-2"))
    assert fetch_calls[-1] is alt_calendar

    backend.get_event(GetEventRequest(id="event-1"))
    assert fetch_calls[-1] is default_calendar


def test_get_event_rejects_event_that_actually_lives_in_a_different_calendar(
    settings, monkeypatch
) -> None:
    """Regression: exchangelib's Account.fetch(folder=...) only uses `folder` to
    validate `only_fields`, not to restrict which item an id resolves to -- an event
    id from a different calendar used to be returned as if it belonged to the
    calendar_id the caller asked for."""
    requested_calendar = SimpleNamespace(id="cal-requested")
    backend = EWSExchangeBackend(settings)

    def fetch(ids, folder=None):
        # The id actually resolves to an event that lives in a different calendar.
        return iter([_fake_event("event-1", False, parent_folder_id="cal-other")])

    backend._account = SimpleNamespace(calendar=SimpleNamespace(id="cal-default"), fetch=fetch)
    monkeypatch.setattr(backend, "_resolve_folder", lambda value: requested_calendar)

    with pytest.raises(APIError) as excinfo:
        backend.get_event(GetEventRequest(id="event-1", calendar_id="cal-requested"))
    assert excinfo.value.code == "not_found"


def test_work_hours_rejects_bad_time_format() -> None:
    with pytest.raises(Exception, match="HH:MM"):
        WorkHours(start="bad", end="18:00")


def test_work_hours_rejects_end_before_start() -> None:
    with pytest.raises(Exception, match="end must be after start"):
        WorkHours(start="18:00", end="09:00")


def test_recurrence_pattern_rejects_invalid_weekday() -> None:
    with pytest.raises(Exception):
        RecurrencePattern(type="weekly", days_of_week=["someday"])


def test_recurrence_pattern_rejects_end_date_and_occurrences_together() -> None:
    with pytest.raises(Exception, match="not both"):
        RecurrencePattern(type="daily", end_date=date(2026, 1, 1), occurrences=5)


def test_recurrence_pattern_rejects_days_of_week_on_non_weekly_type() -> None:
    with pytest.raises(Exception, match="days_of_week"):
        RecurrencePattern(type="daily", days_of_week=["monday"])


def test_recurrence_pattern_rejects_interval_on_yearly_type() -> None:
    """exchangelib's AbsoluteYearlyPattern has no interval field, so any
    interval other than 1 could only ever be silently ignored downstream."""
    with pytest.raises(Exception, match="interval"):
        RecurrencePattern(type="yearly", interval=2)


def test_create_event_rejects_recurrence_end_date_before_start() -> None:
    with pytest.raises(Exception, match="recurrence.end_date"):
        CreateEventRequest.model_validate(
            {
                "subject": "Standup",
                "start": "2026-04-13T09:00:00+00:00",
                "end": "2026-04-13T09:30:00+00:00",
                "recurrence": {"type": "daily", "end_date": "2026-01-01"},
            }
        )


def test_list_calendars_maps_error_raised_while_walking_folders(settings) -> None:
    """Regression guard: list_calendars ran account.root.walk() outside any
    try/except, so an auth/network failure escaped as a raw exchangelib
    exception instead of an APIError."""
    backend = EWSExchangeBackend(settings)

    class PoisonRoot:
        def walk(self):
            raise UnauthorizedError("access denied")

    backend._account = SimpleNamespace(calendar=SimpleNamespace(id="cal-1"), root=PoisonRoot())

    with pytest.raises(AuthFailedError):
        backend.list_calendars()


def test_find_free_slots_never_produces_a_slot_past_end(settings, monkeypatch) -> None:
    """A window that isn't a multiple of the requested duration (100 minutes for
    a 60-minute duration) must yield one full-length slot, not a second slot
    truncated or overflowing past `end`."""
    backend = EWSExchangeBackend(settings)
    start = datetime(2026, 4, 8, 9, 0, tzinfo=UTC)
    end = datetime(2026, 4, 8, 10, 40, tzinfo=UTC)

    class FakeProtocol:
        def get_free_busy_info(self, **kwargs):
            return [SimpleNamespace(merged="0" * 10)]

    backend._account = SimpleNamespace(protocol=FakeProtocol(), default_timezone=UTC)
    monkeypatch.setattr(backend, "_to_ews_datetime", lambda value: value)

    request = FindFreeSlotsRequest(
        attendees=["user@example.com"], duration=60, start=start, end=end
    )
    slots = backend.find_free_slots(request)

    assert len(slots) == 1
    assert slots[0].start == start
    assert slots[0].end == datetime(2026, 4, 8, 10, 0, tzinfo=UTC)
    assert slots[0].end <= end
