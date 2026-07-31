from __future__ import annotations

from datetime import UTC, datetime
from types import SimpleNamespace

import outlook_mcp.exchange_client.calendar as calendar_module
from outlook_mcp.exchange_client import EWSExchangeBackend
from outlook_mcp.models import (
    CreateEventRequest,
    DeleteEventRequest,
    ListEventsRequest,
)


class FakeCalendarItem:
    saved: list["FakeCalendarItem"] = []

    def __init__(self, **kwargs) -> None:
        self.kwargs = kwargs
        self.id = "event-new"

    def save(self, **kwargs) -> None:
        FakeCalendarItem.saved.append(self)


def test_create_event_passes_recurrence_and_online_meeting(settings, monkeypatch) -> None:
    monkeypatch.setattr(calendar_module, "CalendarItem", FakeCalendarItem)
    backend = EWSExchangeBackend(settings)
    backend._account = SimpleNamespace(calendar=object(), default_timezone=UTC)

    request = CreateEventRequest.model_validate(
        {
            "subject": "Sprint planning",
            "start": "2026-04-13T09:00:00+00:00",
            "end": "2026-04-13T10:00:00+00:00",
            "recurrence": {"type": "weekly", "interval": 2, "days_of_week": ["monday", "wednesday"]},
            "online_meeting": True,
        }
    )

    backend.create_event(request)

    saved = FakeCalendarItem.saved[-1]
    assert saved.kwargs["is_online_meeting"] is True
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
    assert saved.kwargs["is_online_meeting"] is False


class FakeCalendarFolder:
    def __init__(self, items: list) -> None:
        self._items = items

    def view(self, start, end):
        return self._items


def _fake_event(item_id: str, is_recurring: bool) -> SimpleNamespace:
    return SimpleNamespace(
        id=item_id,
        subject="Standup",
        start=datetime(2026, 4, 13, 9, 0, tzinfo=UTC),
        end=datetime(2026, 4, 13, 9, 30, tzinfo=UTC),
        is_recurring=is_recurring,
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

    def delete(self, **kwargs) -> None:
        self.deleted_with = kwargs

    def cancel(self, **kwargs) -> None:
        self.cancelled_with = kwargs


def test_delete_event_sends_cancel_message_when_provided(settings) -> None:
    backend = EWSExchangeBackend(settings)
    item = FakeDeletableItem()
    backend._account = SimpleNamespace(
        calendar=object(),
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
        calendar=object(),
        fetch=lambda **kwargs: iter([item]),
    )

    request = DeleteEventRequest.model_validate({"id": "event-1", "notify_attendees": True})

    backend.delete_event(request)

    assert item.deleted_with == {"send_meeting_cancellations": "SendToAllAndSaveCopy"}
    assert item.cancelled_with is None
