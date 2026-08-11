"""The response-diet guarantees: listings fetch and return only what a listing
needs, and the heavy parts (bodies, recipient lists, RFC-822 headers) stay on
the single-item tools that were explicitly asked for them."""

from __future__ import annotations

from datetime import UTC, datetime
from types import SimpleNamespace

import pytest
from exchangelib import CalendarItem, Message
from exchangelib.errors import ErrorItemNotFound, UnauthorizedError
from exchangelib.ewsdatetime import EWSTimeZone

from outlook_mcp.errors import APIError
from outlook_mcp.exchange_client import EWSExchangeBackend
from outlook_mcp.exchange_client.calendar import _EVENT_SUMMARY_FIELDS
from outlook_mcp.exchange_client.email import _EMAIL_SUMMARY_FIELDS, _THREAD_CANDIDATE_FIELDS
from outlook_mcp.models import (
    DeleteEmailRequest,
    EmailFull,
    EmailSummary,
    GetEmailRequest,
    GetEventRequest,
    ListEventsRequest,
)


class _EventQuery(list):
    """folder.view() result: list-like, and it records the projection."""

    projected: tuple = ()

    def only(self, *fields):
        _EventQuery.projected = fields
        return self


def _event_item(**kwargs):
    kwargs.setdefault("id", "event-1")
    kwargs.setdefault("subject", "Sync")
    kwargs.setdefault("start", datetime(2026, 4, 13, 9, 0, tzinfo=UTC))
    kwargs.setdefault("end", datetime(2026, 4, 13, 10, 0, tzinfo=UTC))
    kwargs.setdefault("parent_folder_id", SimpleNamespace(id="cal-1"))
    return SimpleNamespace(**kwargs)


def _calendar_backend(settings, folder=None, item=None) -> EWSExchangeBackend:
    backend = EWSExchangeBackend(settings)
    fields = {
        "default_timezone": EWSTimeZone("UTC"),
        "calendar": folder if folder is not None else SimpleNamespace(id="cal-1"),
    }
    if item is not None:
        fields["fetch"] = lambda **kwargs: iter([item])
    backend._account = SimpleNamespace(**fields)
    return backend


# --- field projections --------------------------------------------------------


@pytest.mark.parametrize(
    ("fields", "item_type"),
    [
        (_EVENT_SUMMARY_FIELDS, CalendarItem),
        (_EMAIL_SUMMARY_FIELDS, Message),
        (_THREAD_CANDIDATE_FIELDS, Message),
    ],
)
def test_every_projected_field_name_exists_in_exchangelib(fields, item_type) -> None:
    """A typo in a projection constant would surface as ErrorInvalidPropertySet
    on a live server only -- the fakes in this suite never validate names."""
    known = {field.name for field in item_type.FIELDS}

    unknown = [name for name in fields if name not in known]

    assert unknown == [], f"unknown {item_type.__name__} fields: {unknown}"


def test_the_event_projection_excludes_the_body() -> None:
    """The body is the reason a week of a busy calendar used to be the heaviest
    response in the server; listings must not fetch it at all."""
    assert "body" not in _EVENT_SUMMARY_FIELDS
    assert "text_body" not in _EVENT_SUMMARY_FIELDS


def test_list_events_projects_the_query_and_omits_bodies(settings) -> None:
    folder = SimpleNamespace(
        id="cal-1",
        view=lambda start, end: _EventQuery([_event_item(body="agenda", text_body="agenda")]),
    )
    backend = _calendar_backend(settings, folder=folder)

    events = backend.list_events(
        ListEventsRequest(
            start=datetime(2026, 4, 13, 0, 0, tzinfo=UTC),
            end=datetime(2026, 4, 14, 0, 0, tzinfo=UTC),
        )
    )

    assert _EventQuery.projected == _EVENT_SUMMARY_FIELDS
    assert events[0].subject == "Sync"
    assert events[0].body is None


def test_get_event_still_returns_the_body(settings) -> None:
    backend = _calendar_backend(settings, item=_event_item(text_body="agenda"))

    event = backend.get_event(GetEventRequest(id="event-1"))

    assert event.body == "agenda"


# --- email summaries and headers ----------------------------------------------


def test_recipient_lists_live_on_the_full_email_not_the_summary() -> None:
    """A 50-row listing of a mailing-list folder would otherwise carry hundreds
    of {name, email} objects nobody asked for."""
    assert "to" not in EmailSummary.model_fields
    assert "to" in EmailFull.model_fields
    assert "to_recipients" not in _EMAIL_SUMMARY_FIELDS


def _mail_item(**kwargs):
    kwargs.setdefault("id", "email-1")
    kwargs.setdefault("subject", "Hello")
    kwargs.setdefault("datetime_received", datetime(2026, 4, 7, 10, 0, tzinfo=UTC))
    kwargs.setdefault("author", SimpleNamespace(email_address="a@example.com", name="A"))
    kwargs.setdefault("to_recipients", [])
    kwargs.setdefault("attachments", [])
    kwargs.setdefault("headers", [SimpleNamespace(name="X-Test", value="1")])
    return SimpleNamespace(**kwargs)


def _mail_backend(settings, item) -> EWSExchangeBackend:
    backend = EWSExchangeBackend(settings)
    backend._account = SimpleNamespace(
        default_timezone=EWSTimeZone("UTC"),
        fetch=lambda **kwargs: iter([item]),
    )
    return backend


def test_get_email_omits_headers_unless_asked(settings) -> None:
    backend = _mail_backend(settings, _mail_item(text_body="Body"))

    result = backend.get_email(GetEmailRequest(id="email-1"))

    assert result.headers == {}


def test_get_email_returns_headers_on_explicit_request(settings) -> None:
    backend = _mail_backend(settings, _mail_item(text_body="Body"))

    result = backend.get_email(GetEmailRequest(id="email-1", include_headers=True))

    assert result.headers == {"X-Test": "1"}


# --- the two-phase thread fetch ------------------------------------------------


def _thin(item_id: str, hour: int) -> SimpleNamespace:
    return SimpleNamespace(
        id=item_id,
        subject="Hello",
        datetime_received=datetime(2026, 4, 7, hour, 0, tzinfo=UTC),
    )


def test_fetch_thread_bodies_refetches_everything_but_the_anchor(settings) -> None:
    anchor = _thin("email-anchor", 9)
    fetched_ids: list[list[str]] = []

    def fetch(ids):
        fetched_ids.append([item_id.id for item_id in ids])
        return iter([SimpleNamespace(id=item_id.id, full=True) for item_id in ids])

    backend = EWSExchangeBackend(settings)
    backend._account = SimpleNamespace(fetch=fetch)

    result = backend._fetch_thread_bodies(
        [_thin("email-1", 8), anchor, _thin("email-2", 10)], anchor
    )

    assert fetched_ids == [["email-1", "email-2"]], "the anchor must not be refetched"
    assert result[1] is anchor
    assert [getattr(item, "full", False) for item in result] == [True, False, True]


def test_fetch_thread_bodies_drops_a_message_deleted_between_the_phases(settings) -> None:
    """ErrorItemNotFound for one id means that message genuinely no longer
    exists; the rest of the thread is still the right answer."""

    def fetch(ids):
        return iter(
            [
                SimpleNamespace(id="email-1", full=True),
                ErrorItemNotFound("gone"),
            ]
        )

    backend = EWSExchangeBackend(settings)
    backend._account = SimpleNamespace(fetch=fetch)

    result = backend._fetch_thread_bodies([_thin("email-1", 8), _thin("email-2", 10)], anchor=None)

    assert [item.id for item in result] == ["email-1"]


def test_fetch_thread_bodies_surfaces_any_other_error(settings) -> None:
    def fetch(ids):
        return iter([UnauthorizedError("bad creds")])

    backend = EWSExchangeBackend(settings)
    backend._account = SimpleNamespace(fetch=fetch)

    with pytest.raises(APIError) as excinfo:
        backend._fetch_thread_bodies([_thin("email-1", 8)], anchor=None)

    assert excinfo.value.code == "auth_failed"


# --- projected single-item fetches for mutations --------------------------------


def test_delete_email_fetches_only_the_folder_field(settings) -> None:
    """delete only invokes a method on the item; downloading the full body and
    attachment metadata first was pure wire cost."""
    captured: dict = {}

    def fetch(**kwargs):
        captured.update(kwargs)
        return iter([SimpleNamespace(id="email-1", move_to_trash=lambda: None)])

    backend = EWSExchangeBackend(settings)
    backend._account = SimpleNamespace(fetch=fetch)

    backend.delete_email(DeleteEmailRequest(id="email-1"))

    assert captured["only_fields"] == ["parent_folder_id"]


def test_fetch_thread_bodies_skips_the_fetch_when_only_the_anchor_survived(settings) -> None:
    """A one-message thread already holds its body; a bulk fetch of zero ids
    would still cost an EWS round trip on some servers."""
    anchor = _thin("email-anchor", 9)
    backend = EWSExchangeBackend(settings)
    backend._account = SimpleNamespace()  # any fetch attempt would blow up here

    result = backend._fetch_thread_bodies([anchor], anchor)

    assert result == [anchor]
