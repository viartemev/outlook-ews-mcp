from __future__ import annotations

from types import SimpleNamespace

from exchangelib.ewsdatetime import EWSDate, EWSDateTime, EWSTimeZone

from outlook_mcp.errors import PermissionDeniedError
from outlook_mcp.exchange_client import EWSExchangeBackend
from outlook_mcp.models import FindFreeSlotsRequest, ListEventsRequest


def _appointment(start, end, subject="Offsite", all_day=False) -> SimpleNamespace:
    return SimpleNamespace(
        id=f"event-{subject}",
        subject=subject,
        start=start,
        end=end,
        location=None,
        organizer=SimpleNamespace(email_address="organizer@example.com", name="Organizer"),
        required_attendees=[],
        optional_attendees=[],
        is_all_day=all_day,
        is_recurring=False,
        my_response_type="Organizer",
        text_body="",
        body="",
        reminder_minutes_before_start=15,
        categories=[],
        recurrence=None,
        importance="Normal",
        meeting_workspace_url=None,
        net_show_url=None,
        legacy_free_busy_status="Busy",
    )


def _backend_with_events(settings, items) -> EWSExchangeBackend:
    backend = EWSExchangeBackend(settings)
    backend._account = SimpleNamespace(
        default_timezone=EWSTimeZone("Europe/Moscow"),
        calendar=SimpleNamespace(view=lambda start, end: _EventQuery(items)),
    )
    return backend


def _two_day_window() -> ListEventsRequest:
    return ListEventsRequest.model_validate(
        {"start": "2026-04-08T00:00:00+00:00", "end": "2026-04-10T00:00:00+00:00"}
    )


def test_all_day_events_keep_a_timezone(settings) -> None:
    # EWS reports all-day appointments with EWSDate, which pydantic turns into a
    # naive midnight datetime, so a day's agenda mixed aware and naive timestamps.
    backend = _backend_with_events(
        settings, [_appointment(EWSDate(2026, 4, 8), EWSDate(2026, 4, 9), all_day=True)]
    )

    event = backend.list_events(_two_day_window())[0]

    assert event.is_all_day is True
    assert event.start.tzinfo is not None
    assert event.end.tzinfo is not None
    assert (event.start.hour, event.start.minute) == (0, 0)


def test_all_day_and_timed_events_sort_together(settings) -> None:
    backend = _backend_with_events(
        settings,
        [
            _appointment(EWSDate(2026, 4, 8), EWSDate(2026, 4, 9), "AllDay", all_day=True),
            _appointment(
                EWSDateTime(2026, 4, 8, 14, 0, tzinfo=EWSTimeZone("UTC")),
                EWSDateTime(2026, 4, 8, 15, 0, tzinfo=EWSTimeZone("UTC")),
                "Sync",
            ),
        ],
    )

    events = backend.list_events(_two_day_window())

    assert sorted(events, key=lambda event: event.start)


def test_availability_survives_an_all_day_event(settings) -> None:
    # _compute_free_slots compares event.start against the requested window, which
    # raised "can't compare offset-naive and offset-aware datetimes".
    backend = _backend_with_events(
        settings, [_appointment(EWSDate(2026, 4, 8), EWSDate(2026, 4, 9), all_day=True)]
    )

    result = backend.get_my_availability(_two_day_window())

    assert result.busy_slots


class _EventQuery(list):
    """folder.view() result: list-like, accepting the .only() projection."""

    def only(self, *fields):
        return self


class FakeCalendarFolder:
    def __init__(self) -> None:
        self.start = None
        self.end = None

    def view(self, start, end):
        self.start = start
        self.end = end
        return _EventQuery()


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
        yield SimpleNamespace(merged="0000")

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
    assert captured["start"].tzinfo.key == "UTC"
    assert captured["end"].tzinfo.key == "UTC"
    assert result[0].start.tzinfo.key == "Europe/Moscow"
    assert result[0].end.tzinfo.key == "Europe/Moscow"


def test_shared_calendar_uses_the_target_mailbox_timezone(settings, monkeypatch) -> None:
    backend = EWSExchangeBackend(settings)
    folder = FakeCalendarFolder()
    target = SimpleNamespace(
        calendar=folder,
        default_timezone=EWSTimeZone("America/New_York"),
    )
    monkeypatch.setattr(backend, "_account_for", lambda mailbox: target)
    monkeypatch.setattr(backend, "_shared_calendar_folder", lambda mailbox: folder)

    backend.list_events(
        ListEventsRequest.model_validate(
            {
                "start": "2026-04-13T09:00:00",
                "end": "2026-04-13T10:00:00",
                "mailbox": "other@example.com",
            }
        )
    )

    assert folder.start.tzinfo.key == "America/New_York"
    assert folder.end.tzinfo.key == "America/New_York"


def test_free_busy_only_share_falls_back_to_availability(settings, monkeypatch) -> None:
    """A calendar shared with free/busy rights only cannot be opened as a folder,
    but the availability service still returns its entries with subject and
    location, the way Outlook shows it."""
    backend = EWSExchangeBackend(settings)
    tz = EWSTimeZone("Europe/Berlin")

    def denied(mailbox):
        raise PermissionDeniedError()

    captured = {}

    def get_free_busy_info(accounts, start, end, requested_view):
        captured.update(accounts=accounts, requested_view=requested_view)
        entry = SimpleNamespace(
            start=EWSDateTime(2026, 4, 13, 9, 0, tzinfo=tz),
            end=EWSDateTime(2026, 4, 13, 10, 0, tzinfo=tz),
            busy_type="Busy",
            details=SimpleNamespace(subject="Team meeting", location="Room 1", is_recurring=True),
        )
        return iter([SimpleNamespace(calendar_events=[entry])])

    account = SimpleNamespace(
        default_timezone=tz,
        protocol=SimpleNamespace(get_free_busy_info=get_free_busy_info),
    )
    monkeypatch.setattr(type(backend), "account", property(lambda self: account))
    monkeypatch.setattr(backend, "_account_for", lambda mailbox: account)
    monkeypatch.setattr(backend, "_shared_calendar_folder", denied)

    request = {
        "start": "2026-04-13T08:00:00",
        "end": "2026-04-13T18:00:00",
        "mailbox": "team@example.com",
    }
    events = backend.list_events(ListEventsRequest.model_validate(request))

    assert captured["accounts"] == [("team@example.com", "Required", False)]
    assert captured["requested_view"] == "Detailed"
    assert [(e.subject, e.location, e.free_busy_status, e.id) for e in events] == [
        ("Team meeting", "Room 1", "busy", "")
    ]
    assert events[0].organizer.email == "team@example.com"

    events = backend.list_events(
        ListEventsRequest.model_validate({**request, "include_recurring": False})
    )
    assert events == []
