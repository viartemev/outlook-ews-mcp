from __future__ import annotations

from datetime import UTC, datetime
from types import SimpleNamespace

import pytest
from exchangelib import HTMLBody
from exchangelib.errors import (
    ErrorAccessDenied,
    ErrorFolderSavePropertyError,
    ResponseMessageError,
    UnauthorizedError,
)
from exchangelib.ewsdatetime import EWSTimeZone

import outlook_mcp.exchange_client.contacts as contacts_module
from outlook_mcp.errors import APIError, NotFoundError
from outlook_mcp.exchange_client import EWSExchangeBackend
from outlook_mcp.models import (
    CreateContactRequest,
    CreateEventRequest,
    DeleteContactRequest,
    DeleteEventRequest,
    GetContactRequest,
    GetEmailRequest,
    ListEventsRequest,
    RecurrencePattern,
    RespondToInviteRequest,
    SearchContactsRequest,
    UpdateContactRequest,
    UpdateEventRequest,
)

_FOLDER = SimpleNamespace(id="folder-1", folder_class="IPF.Appointment")


def _item(**kwargs):
    kwargs.setdefault("id", "item-1")
    kwargs.setdefault("parent_folder_id", SimpleNamespace(id="folder-1"))
    return SimpleNamespace(**kwargs)


def _backend(settings, **fields) -> EWSExchangeBackend:
    fields.setdefault("default_timezone", EWSTimeZone("UTC"))
    fields.setdefault("calendar", _FOLDER)
    backend = EWSExchangeBackend(settings)
    backend._account = SimpleNamespace(**fields)
    return backend


def _raising_fetch(error):
    def fetch(**kwargs):
        raise error

    return fetch


# --- base.py -----------------------------------------------------------------


def test_fetch_item_maps_a_transport_failure(settings) -> None:
    backend = _backend(settings, fetch=_raising_fetch(UnauthorizedError("bad creds")))

    with pytest.raises(APIError) as excinfo:
        backend.get_email(GetEmailRequest(id="email-1"))

    assert excinfo.value.code == "auth_failed"


def test_fetch_item_reports_an_empty_fetch_as_not_found(settings) -> None:
    backend = _backend(settings, fetch=lambda **kwargs: iter([]))

    with pytest.raises(NotFoundError):
        backend.get_email(GetEmailRequest(id="email-1"))


def test_resolve_folder_root_value_returns_root(settings, monkeypatch) -> None:
    import outlook_mcp.exchange_client.base as base_module

    class EmptyIdLookup:
        def __init__(self, account, folders):
            pass

        def resolve(self):
            return iter(())

    monkeypatch.setattr(base_module, "FolderCollection", EmptyIdLookup)
    root = SimpleNamespace(children=[])
    backend = _backend(
        settings,
        root=root,
        archive_root=root,
        inbox=_FOLDER,
        sent=_FOLDER,
        drafts=_FOLDER,
        trash=_FOLDER,
        junk=_FOLDER,
        contacts=_FOLDER,
    )

    assert backend._resolve_folder("root") is root
    # Slashes-only input has no path segments left after splitting.
    assert backend._resolve_folder("//") is root


def test_walk_child_folders_survives_an_unreadable_folder(settings) -> None:
    class Unreadable:
        @property
        def children(self):
            raise ErrorAccessDenied("no access")

    backend = _backend(settings)

    assert backend._walk_child_folders(Unreadable(), ["Sub"]) is None


def test_get_folder_by_id_treats_an_error_element_as_no_match(settings, monkeypatch) -> None:
    """A per-folder error element means "this id did not resolve" and falls
    through to the path lookup; only the GetFolder call itself failing raises."""
    import outlook_mcp.exchange_client.base as base_module

    class ElementError:
        def __init__(self, account, folders):
            pass

        def resolve(self):
            yield ErrorAccessDenied("not yours")

    monkeypatch.setattr(base_module, "FolderCollection", ElementError)
    backend = _backend(settings)
    assert backend._get_folder_by_id("AAA=") is None

    class CallFails:
        def __init__(self, account, folders):
            pass

        def resolve(self):
            raise UnauthorizedError("bad creds")

    monkeypatch.setattr(base_module, "FolderCollection", CallFails)
    with pytest.raises(APIError) as excinfo:
        backend._get_folder_by_id("AAA=")
    assert excinfo.value.code == "auth_failed"


def test_normalize_importance_falls_back_for_unknown_values(settings) -> None:
    backend = _backend(settings)
    assert backend._normalize_importance("Critical") == "normal"
    assert backend._normalize_importance("high") == "high"


def test_extract_message_body_variants(settings) -> None:
    backend = _backend(settings)
    text, html = backend._extract_message_body(
        SimpleNamespace(text_body=None, body=HTMLBody("<p>x</p>"))
    )
    assert text == "<p>x</p>"
    assert html == "<p>x</p>"

    text, html = backend._extract_message_body(SimpleNamespace(text_body=None, body="plain"))
    assert (text, html) == ("plain", None)


def test_headers_are_flattened_and_capped(settings) -> None:
    backend = _backend(settings)
    headers = [
        SimpleNamespace(name="X-One", value="ok"),
        SimpleNamespace(name=None, value="dropped"),
        SimpleNamespace(name="X-Two", value=None),
    ]

    assert backend._headers_to_dict(headers) == {"X-One": "ok"}


def test_to_ews_datetime_localizes_naive_values(settings) -> None:
    backend = _backend(settings)

    converted = backend._to_ews_datetime(datetime(2026, 4, 13, 9, 0))

    assert converted.tzinfo is not None
    assert converted.hour == 9


def test_map_exception_covers_the_remaining_branches(settings) -> None:
    backend = _backend(settings)

    not_found_no_id = backend._map_exception(ResponseMessageError("object not found"))
    assert not_found_no_id.code in {"exchange_error", "not_found"}

    save_error = backend._map_exception(ErrorFolderSavePropertyError("bad folder property"))
    assert save_error.code == "exchange_error"

    timeout = backend._map_exception(TimeoutError("socket timed out"))
    assert timeout.code == "timeout"

    unmapped = backend._map_exception(RuntimeError("who knows"))
    assert unmapped.code == "exchange_unavailable"


def test_ping_maps_an_inbox_failure(settings) -> None:
    class BadInbox:
        @property
        def total_count(self):
            raise UnauthorizedError("bad creds")

    backend = _backend(settings, inbox=BadInbox(), protocol=SimpleNamespace(version=None))

    with pytest.raises(APIError) as excinfo:
        backend.ping()

    assert excinfo.value.code == "auth_failed"


def test_ping_and_mailbox_info_succeed_on_a_healthy_account(settings) -> None:
    backend = _backend(
        settings,
        inbox=SimpleNamespace(total_count=5),
        protocol=SimpleNamespace(version=SimpleNamespace(api_version="Exchange2016")),
        primary_smtp_address="user@example.com",
        fullname="User Example",
    )

    ping = backend.ping()
    assert ping.status == "ok"
    assert ping.version == "Exchange2016"

    info = backend.get_mailbox_info()
    assert info.email_address == "user@example.com"
    assert info.exchange_version == "Exchange2016"


# --- calendar.py -------------------------------------------------------------


def _event_item(**kwargs):
    kwargs.setdefault("id", "event-1")
    kwargs.setdefault("subject", "Sync")
    kwargs.setdefault("start", datetime(2026, 4, 8, 9, 0, tzinfo=UTC))
    kwargs.setdefault("end", datetime(2026, 4, 8, 10, 0, tzinfo=UTC))
    kwargs.setdefault("parent_folder_id", SimpleNamespace(id="folder-1"))
    kwargs.setdefault("organizer", SimpleNamespace(email_address="o@example.com", name="O"))
    kwargs.setdefault("required_attendees", [])
    kwargs.setdefault("optional_attendees", [])
    return SimpleNamespace(**kwargs)


def test_event_moment_passes_datetimes_through(settings) -> None:
    backend = _backend(settings)
    moment = datetime(2026, 4, 8, 9, 0, tzinfo=UTC)
    assert backend._event_moment(moment) is moment
    assert backend._event_moment("not-a-date") == "not-a-date"


def test_attendees_map_both_shapes(settings) -> None:
    backend = _backend(settings)
    wrapped = SimpleNamespace(
        mailbox=SimpleNamespace(email_address="a@example.com", name="A"),
        response_type="Accept",
    )
    bare = SimpleNamespace(email_address="b@example.com", name="B")

    assert backend._to_attendee(wrapped).response_type == "accept"
    assert backend._to_attendee(bare).email == "b@example.com"


def test_list_events_maps_a_view_failure(settings) -> None:
    class _ExplodingIter:
        def __iter__(self):
            raise UnauthorizedError("bad creds")

    class BadView:
        folder_class = "IPF.Appointment"

        def view(self, start, end):
            return _ExplodingIter()

    backend = _backend(settings, calendar=BadView())

    with pytest.raises(APIError) as excinfo:
        backend.list_events(
            ListEventsRequest(start="2026-04-08T00:00:00+00:00", end="2026-04-09T00:00:00+00:00")
        )

    assert excinfo.value.code == "auth_failed"


def test_build_recurrence_covers_every_pattern_type(settings) -> None:
    backend = _backend(settings)
    start = datetime(2026, 4, 8, 9, 0, tzinfo=UTC)

    daily = backend._build_recurrence(RecurrencePattern(type="daily", interval=2), start)
    assert type(daily.pattern).__name__ == "DailyPattern"
    assert daily.pattern.interval == 2

    monthly = backend._build_recurrence(RecurrencePattern(type="monthly", occurrences=3), start)
    assert type(monthly.pattern).__name__ == "AbsoluteMonthlyPattern"
    assert monthly.pattern.day_of_month == 8

    yearly = backend._build_recurrence(
        RecurrencePattern(type="yearly", end_date="2027-01-01"), start
    )
    assert type(yearly.pattern).__name__ == "AbsoluteYearlyPattern"
    assert yearly.pattern.month == 4


def test_create_event_maps_a_save_failure(settings, monkeypatch) -> None:
    import outlook_mcp.exchange_client.calendar as calendar_module

    class ExplodingItem:
        def __init__(self, **kwargs):
            pass

        def save(self, **kwargs):
            raise ErrorAccessDenied("calendar is read-only")

    monkeypatch.setattr(calendar_module, "CalendarItem", ExplodingItem)
    backend = _backend(settings)

    with pytest.raises(APIError) as excinfo:
        backend.create_event(
            CreateEventRequest(
                subject="S", start="2026-04-08T09:00:00+00:00", end="2026-04-08T10:00:00+00:00"
            )
        )

    assert excinfo.value.code == "permission_denied"


def test_update_event_floors_all_day_bounds_and_maps_save_failures(settings) -> None:
    item = _event_item(is_all_day=True)
    item.save = lambda **kwargs: (_ for _ in ()).throw(ErrorAccessDenied("locked"))
    backend = _backend(settings, fetch=lambda **kwargs: iter([item]))

    with pytest.raises(APIError) as excinfo:
        backend.update_event(UpdateEventRequest(id="event-1", start="2026-04-08T15:30:00+00:00"))

    assert excinfo.value.code == "permission_denied"
    # The all-day start was floored to midnight before the save was attempted.
    assert item.start.hour == 0
    assert item.start.day == 8


def test_delete_event_maps_a_failure(settings) -> None:
    item = _event_item()
    item.delete = lambda **kwargs: (_ for _ in ()).throw(ErrorAccessDenied("locked"))
    backend = _backend(settings, fetch=lambda **kwargs: iter([item]))

    with pytest.raises(APIError):
        backend.delete_event(DeleteEventRequest(id="event-1", notify_attendees=False))


@pytest.mark.parametrize(
    ("response", "method_name"),
    [("accept", "accept"), ("tentative", "tentatively_accept"), ("decline", "decline")],
)
def test_respond_to_invite_calls_the_matching_method(settings, response, method_name) -> None:
    calls: list[str] = []
    item = _event_item()
    for name in ("accept", "tentatively_accept", "decline"):
        setattr(item, name, lambda body=None, _n=name: calls.append(_n))
    backend = _backend(settings, fetch=lambda **kwargs: iter([item]))

    result = backend.respond_to_invite(RespondToInviteRequest(id="event-1", response=response))

    assert calls == [method_name]
    assert result.status == response


def test_respond_to_invite_maps_a_failure(settings) -> None:
    item = _event_item()
    item.accept = lambda body=None: (_ for _ in ()).throw(ErrorAccessDenied("nope"))
    backend = _backend(settings, fetch=lambda **kwargs: iter([item]))

    with pytest.raises(APIError):
        backend.respond_to_invite(RespondToInviteRequest(id="event-1", response="accept"))


def test_get_my_availability_maps_a_listing_failure(settings) -> None:
    class _ExplodingIter:
        def __iter__(self):
            raise UnauthorizedError("bad creds")

    class BadView:
        folder_class = "IPF.Appointment"

        def view(self, start, end):
            return _ExplodingIter()

    backend = _backend(settings, calendar=BadView())

    with pytest.raises(APIError):
        backend.get_my_availability(
            ListEventsRequest(start="2026-04-08T00:00:00+00:00", end="2026-04-09T00:00:00+00:00")
        )


def test_list_calendars_reports_the_default_calendar(settings) -> None:
    calendar = SimpleNamespace(
        id="cal-1", name="Календарь", folder_class="IPF.Appointment", children=[]
    )
    root = SimpleNamespace(walk=lambda: iter([calendar]))
    backend = _backend(
        settings, calendar=calendar, root=root, primary_smtp_address="user@example.com"
    )

    result = backend.list_calendars()

    assert result[0].id == "cal-1"
    assert result[0].is_default is True


# --- contacts.py -------------------------------------------------------------


def _stub_resolve_names(monkeypatch, entries) -> None:
    class FakeResolveNames:
        def __init__(self, protocol) -> None:
            pass

        def call(self, **kwargs):
            if isinstance(entries, Exception):
                raise entries
            yield from entries

    monkeypatch.setattr(contacts_module, "ResolveNames", FakeResolveNames)


def test_personal_search_maps_a_folder_failure(settings) -> None:
    class BadContacts:
        def filter(self, **kwargs):
            raise UnauthorizedError("bad creds")

    backend = _backend(settings, contacts=BadContacts())

    with pytest.raises(APIError):
        backend.search_contacts(SearchContactsRequest(query="ivan", source="personal"))


def test_gal_search_maps_a_resolve_failure(settings, monkeypatch) -> None:
    _stub_resolve_names(monkeypatch, UnauthorizedError("bad creds"))
    backend = _backend(settings, protocol=object())

    with pytest.raises(APIError):
        backend.search_contacts(SearchContactsRequest(query="ivan", source="gal"))


def test_gal_search_skips_mailboxes_without_smtp_and_honours_the_limit(
    settings, monkeypatch
) -> None:
    x500 = SimpleNamespace(email_address="/o=ORG/cn=one", name="No Smtp")
    smtp = SimpleNamespace(email_address="two@example.com", name="Has Smtp")
    _stub_resolve_names(monkeypatch, [(x500, None), (smtp, None), (smtp, None)])
    backend = _backend(settings, protocol=object())

    result = backend.search_contacts(SearchContactsRequest(query="x", source="gal", limit=2))

    assert [c.display_name for c in result] == ["Has Smtp", "Has Smtp"]


def test_gal_lookup_maps_error_entries(settings, monkeypatch) -> None:
    _stub_resolve_names(monkeypatch, [ErrorAccessDenied("gal is closed")])
    backend = _backend(settings, protocol=object())

    with pytest.raises(APIError) as excinfo:
        backend.get_contact(GetContactRequest(id="someone@example.com"))

    assert excinfo.value.code == "permission_denied"


def test_gal_lookup_builds_a_minimal_contact_from_a_bare_mailbox(settings, monkeypatch) -> None:
    mailbox = SimpleNamespace(email_address="smtp:person@example.com", name="Person")
    _stub_resolve_names(monkeypatch, [(mailbox, None)])
    backend = _backend(settings, protocol=object())

    contact = backend.get_contact(GetContactRequest(id="person@example.com"))

    assert contact.display_name == "Person"
    assert contact.email_addresses[0].address == "person@example.com"


def test_create_contact_maps_a_save_failure(settings, monkeypatch) -> None:
    class ExplodingContact:
        def __init__(self, **kwargs):
            pass

        def save(self):
            raise ErrorAccessDenied("contacts are read-only")

    monkeypatch.setattr(contacts_module, "Contact", ExplodingContact)
    backend = _backend(settings, contacts=SimpleNamespace())

    with pytest.raises(APIError):
        backend.create_contact(CreateContactRequest(display_name="Ivan"))


def test_update_contact_maps_a_save_failure(settings) -> None:
    item = _item(
        display_name="Ivan",
        email_addresses=[],
        phone_numbers=[],
    )
    item.save = lambda **kwargs: (_ for _ in ()).throw(ErrorAccessDenied("locked"))
    backend = _backend(
        settings, contacts=SimpleNamespace(id="folder-1"), fetch=lambda **kwargs: iter([item])
    )

    with pytest.raises(APIError):
        backend.update_contact(UpdateContactRequest(id="contact-1", display_name="New"))


def test_delete_contact_maps_a_failure(settings) -> None:
    item = _item()
    item.move_to_trash = lambda: (_ for _ in ()).throw(ErrorAccessDenied("locked"))
    backend = _backend(
        settings, contacts=SimpleNamespace(id="folder-1"), fetch=lambda **kwargs: iter([item])
    )

    with pytest.raises(APIError):
        backend.delete_contact(DeleteContactRequest(id="contact-1"))
