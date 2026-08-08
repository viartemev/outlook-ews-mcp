from __future__ import annotations

import os
import threading
from pathlib import Path
from types import SimpleNamespace

import pytest
from exchangelib.errors import (
    ErrorAccessDenied,
    ErrorInvalidIdMalformed,
    ErrorIrresolvableConflict,
    ErrorItemNotFound,
    ErrorItemSavePropertyError,
    ErrorServerBusy,
    ErrorTimeoutExpired,
    RateLimitError,
    UnauthorizedError,
)
from exchangelib.ewsdatetime import EWSTimeZone
from exchangelib.folders import Inbox
from exchangelib import CalendarItem, Message
from exchangelib.restriction import Q, Restriction
from exchangelib.version import EXCHANGE_2016, Version

from exchangelib.items import HARD_DELETE, MOVE_TO_DELETED_ITEMS

import outlook_mcp.exchange_client.base as exchange_client_base
import outlook_mcp.exchange_client.contacts as exchange_client_contacts
import outlook_mcp.exchange_client.email as exchange_client_email
from outlook_mcp.errors import APIError, AuthFailedError, NotFoundError
from outlook_mcp.exchange_client import EWSExchangeBackend
from outlook_mcp.models import (
    ActionResult,
    DeleteContactRequest,
    DeleteFolderRequest,
    FindFreeSlotsRequest,
    FolderActionRequest,
    ForwardEmailRequest,
    GetContactRequest,
    GetEmailRequest,
    ListEmailsRequest,
    RenameFolderRequest,
    ReplyEmailRequest,
    SearchContactsRequest,
    SearchEmailsRequest,
    SendResult,
)


def _free_slots_request() -> FindFreeSlotsRequest:
    return FindFreeSlotsRequest.model_validate(
        {
            "attendees": ["user@example.com"],
            "duration": 60,
            "start": "2026-04-13T09:00:00+03:00",
            "end": "2026-04-13T11:00:00+03:00",
        }
    )


def _fake_account(views, captured: dict | None = None) -> SimpleNamespace:
    """An account whose protocol answers get_free_busy_info with a generator, like exchangelib does."""

    def get_free_busy_info(**kwargs):
        if captured is not None:
            captured.update(kwargs)
        yield from views

    return SimpleNamespace(
        default_timezone=EWSTimeZone("Europe/Moscow"),
        protocol=SimpleNamespace(get_free_busy_info=get_free_busy_info),
    )


@pytest.mark.parametrize(
    ("configured", "expected"),
    [
        ("mail.example.com", "https://mail.example.com/EWS/Exchange.asmx"),
        ("https://mail.example.com", "https://mail.example.com/EWS/Exchange.asmx"),
        ("https://mail.example.com/", "https://mail.example.com/EWS/Exchange.asmx"),
        (
            "https://mail.example.com/EWS/Exchange.asmx",
            "https://mail.example.com/EWS/Exchange.asmx",
        ),
        (
            "https://mail.example.com/ews/exchange.asmx",
            "https://mail.example.com/ews/exchange.asmx",
        ),
        # A URL copied from a browser keeps its trailing slash. That used to miss
        # the suffix check and get a second "/EWS/Exchange.asmx" appended, leaving
        # a silently unreachable endpoint.
        (
            "https://mail.example.com/EWS/Exchange.asmx/",
            "https://mail.example.com/EWS/Exchange.asmx",
        ),
        (
            "  https://mail.example.com/EWS/Exchange.asmx  ",
            "https://mail.example.com/EWS/Exchange.asmx",
        ),
        ("http://mail.example.com", "http://mail.example.com/EWS/Exchange.asmx"),
    ],
)
def test_normalize_service_endpoint(settings, configured, expected) -> None:
    backend = EWSExchangeBackend(settings)

    assert backend._normalize_service_endpoint(configured) == expected


def test_normalize_service_endpoint_never_repeats_the_ews_path(settings) -> None:
    backend = EWSExchangeBackend(settings)

    for configured in (
        "mail.example.com",
        "https://mail.example.com/EWS/Exchange.asmx",
        "https://mail.example.com/EWS/Exchange.asmx/",
    ):
        assert backend._normalize_service_endpoint(configured).count("Exchange.asmx") == 1


def test_map_exception_classifies_by_exchangelib_type_not_message_text(settings) -> None:
    """Each of these must be matched by its actual exchangelib exception class, not
    by substring-matching the (English, server-supplied) exception message."""
    backend = EWSExchangeBackend(settings)

    conflict = backend._map_exception(ErrorIrresolvableConflict("conflict, retry"))
    assert conflict.code == "conflict"

    timeout = backend._map_exception(ErrorTimeoutExpired("the operation timed out"))
    assert timeout.code == "timeout"

    not_found = backend._map_exception(ErrorItemNotFound("no such item"), item_id="item-1")
    assert not_found.code == "not_found"
    assert not_found.extra["id"] == "item-1"

    # An invalid/read-only property in the save request is not a concurrent-edit
    # conflict and must not be reported as one.
    save_property = backend._map_exception(ErrorItemSavePropertyError("bad property"))
    assert save_property.code == "exchange_error"

    permission = backend._map_exception(ErrorAccessDenied("localized server message"))
    assert permission.code == "permission_denied"


def test_map_exception_marks_transient_busy_errors_as_retryable(settings) -> None:
    """These must be safe for ExchangeClient._retry_read to retry automatically on
    read-only calls -- the account's retry_policy is FailFast, so exchangelib no
    longer retries them internally (see exchange_client/base.py)."""
    backend = EWSExchangeBackend(settings)

    busy = backend._map_exception(ErrorServerBusy("server is busy"))
    assert busy.retryable is True

    timeout = backend._map_exception(ErrorTimeoutExpired("the operation timed out"))
    assert timeout.retryable is True

    rate_limited = backend._map_exception(RateLimitError("max timeout reached", wait=5))
    assert rate_limited.retryable is True


def test_map_exception_leaves_non_transient_errors_not_retryable(settings) -> None:
    backend = EWSExchangeBackend(settings)

    conflict = backend._map_exception(ErrorIrresolvableConflict("conflict, retry"))
    assert conflict.retryable is False

    not_found = backend._map_exception(ErrorItemNotFound("no such item"), item_id="item-1")
    assert not_found.retryable is False

    permission = backend._map_exception(ErrorAccessDenied("access denied"))
    assert permission.retryable is False


def test_map_exception_does_not_log_raw_exchange_message(settings, caplog) -> None:
    backend = EWSExchangeBackend(settings)
    secret_detail = "mail.internal.example user@example.com"

    backend._map_exception(ErrorInvalidIdMalformed(secret_detail))

    assert secret_detail not in caplog.text


def test_find_free_slots_accepts_generator_response(settings) -> None:
    backend = EWSExchangeBackend(settings)
    captured: dict = {}
    backend._account = _fake_account([SimpleNamespace(merged="00")], captured)

    slots = backend.find_free_slots(_free_slots_request())

    assert captured["accounts"] == [("user@example.com", "Required", False)]
    assert captured["merged_free_busy_interval"] == 60
    assert len(slots) == 2
    assert slots[0].start.hour == 9
    assert slots[1].start.hour == 10


def test_find_free_slots_returns_empty_on_generator_without_views(settings) -> None:
    backend = EWSExchangeBackend(settings)
    backend._account = _fake_account([])

    assert backend.find_free_slots(_free_slots_request()) == []


def test_find_free_slots_raises_when_exchange_returns_error(settings) -> None:
    backend = EWSExchangeBackend(settings)
    backend._account = _fake_account([ErrorInvalidIdMalformed("Id is malformed.")])

    with pytest.raises(APIError) as excinfo:
        backend.find_free_slots(_free_slots_request())

    assert excinfo.value.code == "exchange_error"


def test_find_free_slots_requires_every_attendee_free(settings) -> None:
    """A slot free for the first attendee but busy for a later one must not be reported."""
    backend = EWSExchangeBackend(settings)
    request = FindFreeSlotsRequest.model_validate(
        {
            "attendees": ["free@example.com", "busy@example.com"],
            "duration": 60,
            "start": "2026-04-13T09:00:00+03:00",
            "end": "2026-04-13T11:00:00+03:00",
        }
    )
    backend._account = _fake_account([SimpleNamespace(merged="00"), SimpleNamespace(merged="02")])

    slots = backend.find_free_slots(request)

    assert len(slots) == 1
    assert slots[0].start.hour == 9


def test_find_free_slots_clamps_sub_5_minute_duration_to_ews_minimum(settings) -> None:
    """EWS rejects merged_free_busy_interval below 5, so a short meeting duration
    (e.g. 2 minutes) must still request a valid 5-minute sample, and slots must be
    matched against the sample that covers each short slot's start time."""
    backend = EWSExchangeBackend(settings)
    captured: dict = {}
    request = FindFreeSlotsRequest.model_validate(
        {
            "attendees": ["user@example.com"],
            "duration": 2,
            "start": "2026-04-13T09:00:00+03:00",
            "end": "2026-04-13T09:10:00+03:00",
        }
    )
    # Two 5-minute samples: first free ("0"), second busy ("1").
    backend._account = _fake_account([SimpleNamespace(merged="01")], captured)

    slots = backend.find_free_slots(request)

    assert captured["merged_free_busy_interval"] == 5
    assert [slot.start.minute for slot in slots] == [0, 2, 4]


def test_find_free_slots_filters_by_work_hours(settings) -> None:
    backend = EWSExchangeBackend(settings)
    request = FindFreeSlotsRequest.model_validate(
        {
            "attendees": ["user@example.com"],
            "duration": 60,
            "start": "2026-04-13T08:00:00+03:00",
            "end": "2026-04-13T11:00:00+03:00",
            "work_hours": {"start": "09:00", "end": "18:00"},
        }
    )
    backend._account = _fake_account([SimpleNamespace(merged="000")])

    slots = backend.find_free_slots(request)

    assert len(slots) == 2
    assert slots[0].start.hour == 9
    assert slots[1].start.hour == 10


def test_fetch_item_raises_api_error_when_exchange_returns_error(settings) -> None:
    backend = EWSExchangeBackend(settings)

    def fetch(ids, folder=None):
        yield ErrorInvalidIdMalformed("Id is malformed.")

    backend._account = SimpleNamespace(fetch=fetch)

    with pytest.raises(APIError) as excinfo:
        backend.get_email(GetEmailRequest(id="not-an-ews-id"))

    assert excinfo.value.code == "exchange_error"


def test_fetch_item_rejects_item_whose_parent_folder_does_not_match_requested_folder(
    settings,
) -> None:
    """Account.fetch(folder=...) only uses `folder` to validate `only_fields` against
    that folder's allowed fields -- it does NOT restrict which item an id resolves to.
    Passing a calendar_id/folder as a scoping argument must not let an id belonging to
    a different folder (e.g. a different calendar, or a personal contact when a GAL
    lookup was intended) be fetched and acted on as if it lived in the requested one."""
    backend = EWSExchangeBackend(settings)
    item = SimpleNamespace(id="item-1", parent_folder_id=SimpleNamespace(id="folder-other"))
    backend._account = SimpleNamespace(fetch=lambda **kwargs: iter([item]))
    requested_folder = SimpleNamespace(id="folder-requested")

    with pytest.raises(NotFoundError):
        backend._fetch_item("item-1", folder=requested_folder)


def test_fetch_item_accepts_item_whose_parent_folder_matches_requested_folder(settings) -> None:
    backend = EWSExchangeBackend(settings)
    item = SimpleNamespace(id="item-1", parent_folder_id=SimpleNamespace(id="folder-1"))
    backend._account = SimpleNamespace(fetch=lambda **kwargs: iter([item]))
    requested_folder = SimpleNamespace(id="folder-1")

    assert backend._fetch_item("item-1", folder=requested_folder) is item


def test_fetch_item_rejects_item_with_no_resolvable_parent_folder_when_scoped(settings) -> None:
    """A missing parent_folder_id (or an unresolved folder id on the requested folder)
    must fail closed rather than silently skipping the scoping check."""
    backend = EWSExchangeBackend(settings)
    item = SimpleNamespace(id="item-1", parent_folder_id=None)
    backend._account = SimpleNamespace(fetch=lambda **kwargs: iter([item]))
    requested_folder = SimpleNamespace(id="folder-1")

    with pytest.raises(NotFoundError):
        backend._fetch_item("item-1", folder=requested_folder)


def test_fetch_item_rejects_wrong_item_type(settings) -> None:
    backend = EWSExchangeBackend(settings)
    event = CalendarItem(id="event-1")
    backend._account = SimpleNamespace(fetch=lambda **kwargs: iter([event]))

    with pytest.raises(APIError) as excinfo:
        backend._fetch_item("event-1", expected_type=Message)

    assert excinfo.value.code == "not_found"


def test_get_contact_reads_notes_from_body_not_readonly_notes_field(settings) -> None:
    """Regression test: Contact.notes is read-only in exchangelib, so reading it
    always returned None. Outlook stores contact notes in the item body."""
    backend = EWSExchangeBackend(settings)
    item = SimpleNamespace(
        id="contact-1",
        display_name="Ivan Ivanov",
        file_as=None,
        given_name=None,
        surname=None,
        email_addresses=[],
        phone_numbers=[],
        physical_addresses=[],
        company_name=None,
        job_title=None,
        department=None,
        manager=None,
        birthday=None,
        text_body="Met at the conference",
        parent_folder_id=SimpleNamespace(id="contacts-1"),
    )
    backend._account = SimpleNamespace(
        fetch=lambda **kwargs: iter([item]), contacts=SimpleNamespace(id="contacts-1")
    )

    contact = backend.get_contact(GetContactRequest(id="contact-1"))

    assert contact.notes == "Met at the conference"


def test_get_contact_rejects_item_that_actually_lives_in_a_different_folder(settings) -> None:
    """Regression: an id resolving to an item outside the personal Contacts folder
    (e.g. moved to another folder, or belonging to a different item type entirely)
    must not be returned as if it were a contact in Contacts -- Account.fetch(folder=...)
    doesn't enforce that on its own."""
    backend = EWSExchangeBackend(settings)
    item = SimpleNamespace(id="contact-1", parent_folder_id=SimpleNamespace(id="some-other-folder"))
    backend._account = SimpleNamespace(
        fetch=lambda **kwargs: iter([item]), contacts=SimpleNamespace(id="contacts-1")
    )

    with pytest.raises(APIError) as excinfo:
        backend.get_contact(GetContactRequest(id="contact-1"))
    assert excinfo.value.code == "not_found"


def test_get_contact_resolves_gal_address(settings, monkeypatch) -> None:
    backend = EWSExchangeBackend(settings)
    captured: dict = {}
    gal_contact = SimpleNamespace(
        id=None,
        display_name="Ivan Ivanov",
        file_as=None,
        given_name="Ivan",
        surname="Ivanov",
        email_addresses=[
            SimpleNamespace(
                label="EmailAddress1",
                email="X500:/o=ORG/ou=Exchange Administrative Group/cn=ivan",
            ),
            SimpleNamespace(label="EmailAddress2", email="SMTP:ivan@example.com"),
        ],
        phone_numbers=[SimpleNamespace(label="BusinessPhone", phone_number="+79990000000")],
        physical_addresses=[],
        company_name="Example",
        job_title="Manager",
        department="Sales",
        manager=None,
        notes=None,
        birthday=None,
    )

    class FakeResolveNames:
        def __init__(self, protocol) -> None:
            self.protocol = protocol

        def call(self, **kwargs):
            captured.update(kwargs)
            yield SimpleNamespace(email_address="ivan@example.com", name="Ivan Ivanov"), gal_contact

    def fetch(ids, folder=None):
        raise AssertionError("GAL contacts must not be fetched from the personal contacts folder")

    monkeypatch.setattr(exchange_client_contacts, "ResolveNames", FakeResolveNames)
    backend._account = SimpleNamespace(fetch=fetch, contacts=object(), protocol=object())

    contact = backend.get_contact(GetContactRequest(id="ivan@example.com"))

    assert captured["unresolved_entries"] == ["ivan@example.com"]
    assert contact.source == "gal"
    assert contact.id == "ivan@example.com"
    assert contact.display_name == "Ivan Ivanov"
    assert [entry.address for entry in contact.email_addresses] == ["ivan@example.com"]
    assert [entry.number for entry in contact.phone_numbers] == ["+79990000000"]
    assert contact.job_title == "Manager"


def test_get_contact_keeps_only_smtp_addresses_of_gal_contact(settings, monkeypatch) -> None:
    """The GAL returns X500 proxy addresses that are not valid SMTP addresses."""
    backend = EWSExchangeBackend(settings)
    gal_contact = SimpleNamespace(
        display_name="Ivan Ivanov",
        file_as=None,
        email_addresses=[
            SimpleNamespace(
                label="EmailAddress1",
                email="X500:/o=TANDER/ou=Exchange Administrative Group/cn=ivan",
            ),
            SimpleNamespace(label="EmailAddress2", email="SMTP:ivan.alias@example.com"),
        ],
        phone_numbers=[SimpleNamespace(label="BusinessPhone", phone_number=None)],
        physical_addresses=[],
    )

    class FakeResolveNames:
        def __init__(self, protocol) -> None:
            self.protocol = protocol

        def call(self, **kwargs):
            yield SimpleNamespace(email_address="ivan@example.com", name="Ivan Ivanov"), gal_contact

    monkeypatch.setattr(exchange_client_contacts, "ResolveNames", FakeResolveNames)
    backend._account = SimpleNamespace(contacts=object(), protocol=object())

    contact = backend.get_contact(GetContactRequest(id="ivan@example.com"))

    assert [entry.address for entry in contact.email_addresses] == [
        "ivan@example.com",
        "ivan.alias@example.com",
    ]
    assert contact.phone_numbers == []


def test_get_contact_raises_not_found_when_gal_has_no_match(settings, monkeypatch) -> None:
    backend = EWSExchangeBackend(settings)

    class FakeResolveNames:
        def __init__(self, protocol) -> None:
            self.protocol = protocol

        def call(self, **kwargs):
            return iter(())

    monkeypatch.setattr(exchange_client_contacts, "ResolveNames", FakeResolveNames)
    backend._account = SimpleNamespace(contacts=object(), protocol=object())

    with pytest.raises(APIError) as excinfo:
        backend.get_contact(GetContactRequest(id="nobody@example.com"))

    assert excinfo.value.code == "not_found"


def test_search_contacts_matches_personal_contact_by_email_not_just_display_name(
    settings, monkeypatch
) -> None:
    """Regression guard: search_contacts used to filter personal contacts on
    display_name only, so a query matching only the email/company/job_title
    found nothing."""
    backend = EWSExchangeBackend(settings)
    matched_contact = SimpleNamespace(
        id="contact-1",
        display_name="Ivan Ivanov",
        file_as=None,
        email_addresses=[SimpleNamespace(email="ivan@example.com")],
        phone_numbers=[],
        company_name="Acme",
        job_title="Manager",
        department=None,
    )
    captured: dict = {}

    class FakeQuerySet(list):
        def only(self, *fields):
            return self

    class FakeContactsFolder:
        def filter(self, restriction):
            captured["restriction"] = restriction
            return FakeQuerySet([matched_contact])

    backend._account = SimpleNamespace(contacts=FakeContactsFolder(), protocol=object())

    results = backend.search_contacts(
        SearchContactsRequest(query="ivan@example.com", source="personal")
    )

    assert [r.id for r in results] == ["contact-1"]
    # The filter must actually search more than display_name.
    assert "email_addresses" in str(captured["restriction"])


def test_search_contacts_propagates_resolve_names_error_instead_of_unpacking_it(
    settings, monkeypatch
) -> None:
    """Regression guard: an Exception entry yielded by ResolveNames used to be
    unpacked directly as `(mailbox, contact)`, which raises a confusing
    unpacking TypeError instead of the mapped APIError."""
    backend = EWSExchangeBackend(settings)

    class FakeResolveNames:
        def __init__(self, protocol) -> None:
            self.protocol = protocol

        def call(self, **kwargs):
            yield UnauthorizedError("access denied")

    monkeypatch.setattr(exchange_client_contacts, "ResolveNames", FakeResolveNames)
    backend._account = SimpleNamespace(contacts=object(), protocol=object())

    with pytest.raises(AuthFailedError):
        backend.search_contacts(SearchContactsRequest(query="ivan", source="gal"))


def test_search_contacts_gal_result_id_is_resolvable_by_get_contact(settings, monkeypatch) -> None:
    """Regression guard: search_contacts(source="gal") used to return the GAL
    contact's opaque contact.id, which get_contact() cannot resolve (its "@"
    heuristic routes it to the personal contacts folder instead of the GAL).
    The id search_contacts hands back must be the SMTP address so the two
    calls round-trip."""
    backend = EWSExchangeBackend(settings)
    gal_contact = SimpleNamespace(
        id="AAA=opaque-ad-object-id",
        display_name="Ivan Ivanov",
        file_as=None,
        email_addresses=[
            SimpleNamespace(
                label="EmailAddress1",
                email="X500:/o=ORG/ou=Exchange Administrative Group/cn=ivan",
            ),
            SimpleNamespace(label="EmailAddress2", email="SMTP:ivan@example.com"),
        ],
        phone_numbers=[],
        company_name=None,
        job_title=None,
        department=None,
    )

    class FakeResolveNames:
        def __init__(self, protocol) -> None:
            self.protocol = protocol

        def call(self, **kwargs):
            yield SimpleNamespace(email_address="ivan@example.com", name="Ivan Ivanov"), gal_contact

    monkeypatch.setattr(exchange_client_contacts, "ResolveNames", FakeResolveNames)
    backend._account = SimpleNamespace(contacts=object(), protocol=object())

    results = backend.search_contacts(SearchContactsRequest(query="ivan", source="gal"))

    assert len(results) == 1
    assert results[0].id == "ivan@example.com"
    assert results[0].email_addresses == ["ivan@example.com"]

    def fetch(ids, folder=None):
        raise AssertionError("GAL contacts must not be fetched from the personal contacts folder")

    backend._account = SimpleNamespace(fetch=fetch, contacts=object(), protocol=object())
    contact = backend.get_contact(GetContactRequest(id=results[0].id))

    assert contact.source == "gal"
    assert contact.display_name == "Ivan Ivanov"


def test_search_contacts_skips_gal_result_without_smtp_address(settings, monkeypatch) -> None:
    backend = EWSExchangeBackend(settings)
    gal_contact = SimpleNamespace(
        id="AAA=opaque-ad-object-id",
        display_name="Legacy Recipient",
        file_as=None,
        email_addresses=[
            SimpleNamespace(
                label="EmailAddress1",
                email="X500:/o=ORG/ou=Exchange Administrative Group/cn=legacy",
            )
        ],
        phone_numbers=[],
        company_name=None,
        job_title=None,
        department=None,
    )

    class FakeResolveNames:
        def __init__(self, protocol) -> None:
            self.protocol = protocol

        def call(self, **kwargs):
            yield SimpleNamespace(email_address="/o=ORG/ou=legacy"), gal_contact

    monkeypatch.setattr(exchange_client_contacts, "ResolveNames", FakeResolveNames)
    backend._account = SimpleNamespace(contacts=object(), protocol=object())

    assert backend.search_contacts(SearchContactsRequest(query="legacy", source="gal")) == []


def test_get_contact_explicit_source_bypasses_at_sign_heuristic(settings, monkeypatch) -> None:
    """A legacy Exchange DN (no "@") passed with an explicit source="gal" must
    be resolved against the GAL instead of being misrouted to the personal
    contacts folder by the "@"-based heuristic."""
    backend = EWSExchangeBackend(settings)
    legacy_dn = "/o=ExchangeLabs/ou=Exchange Administrative Group/cn=Recipients/cn=ivan"
    gal_contact = SimpleNamespace(
        id=None,
        display_name="Ivan Ivanov",
        file_as=None,
        email_addresses=[SimpleNamespace(label="EmailAddress1", email="ivan@example.com")],
        phone_numbers=[],
        company_name=None,
        job_title=None,
        department=None,
        physical_addresses=[],
        manager=None,
        notes=None,
        birthday=None,
    )

    class FakeResolveNames:
        def __init__(self, protocol) -> None:
            self.protocol = protocol

        def call(self, **kwargs):
            yield SimpleNamespace(email_address="ivan@example.com", name="Ivan Ivanov"), gal_contact

    def fetch(ids, folder=None):
        raise AssertionError(
            "Legacy DN with source='gal' must not hit the personal contacts folder"
        )

    monkeypatch.setattr(exchange_client_contacts, "ResolveNames", FakeResolveNames)
    backend._account = SimpleNamespace(fetch=fetch, contacts=object(), protocol=object())

    contact = backend.get_contact(GetContactRequest(id=legacy_dn, source="gal"))

    assert contact.source == "gal"
    assert contact.display_name == "Ivan Ivanov"


def test_delete_contact_soft_deletes_by_default(settings) -> None:
    backend = EWSExchangeBackend(settings)
    contacts_folder = SimpleNamespace(id="contacts-folder")
    calls: list[str] = []
    contact = SimpleNamespace(
        id="contact-1",
        parent_folder_id=SimpleNamespace(id="contacts-folder"),
        move_to_trash=lambda: calls.append("move_to_trash"),
        delete=lambda: calls.append("delete"),
    )
    backend._account = SimpleNamespace(
        fetch=lambda **kwargs: iter([contact]), contacts=contacts_folder
    )

    result = backend.delete_contact(DeleteContactRequest(id="contact-1"))

    assert calls == ["move_to_trash"]
    assert result.status == "deleted"


def test_delete_contact_hard_delete_bypasses_trash(settings) -> None:
    """Mirrors delete_email's hard_delete option so callers can permanently
    remove a contact instead of just moving it to Deleted Items."""
    backend = EWSExchangeBackend(settings)
    contacts_folder = SimpleNamespace(id="contacts-folder")
    calls: list[str] = []
    contact = SimpleNamespace(
        id="contact-1",
        parent_folder_id=SimpleNamespace(id="contacts-folder"),
        move_to_trash=lambda: calls.append("move_to_trash"),
        delete=lambda: calls.append("delete"),
    )
    backend._account = SimpleNamespace(
        fetch=lambda **kwargs: iter([contact]), contacts=contacts_folder
    )

    result = backend.delete_contact(DeleteContactRequest(id="contact-1", hard_delete=True))

    assert calls == ["delete"]
    assert result.status == "deleted"


def _fake_account_for_folder_resolution(root) -> SimpleNamespace:
    return SimpleNamespace(
        root=root,
        inbox=object(),
        sent=object(),
        drafts=object(),
        trash=object(),
        junk=object(),
        calendar=object(),
        contacts=object(),
    )


def test_resolve_folder_by_id_uses_targeted_lookup_not_full_walk(settings, monkeypatch) -> None:
    backend = EWSExchangeBackend(settings)
    resolved_folder = SimpleNamespace(id="AAA=", name="Projects")
    captured: dict = {}

    class FakeFolderCollection:
        def __init__(self, account, folders) -> None:
            captured["folders"] = folders

        def resolve(self):
            yield resolved_folder

    class PoisonRoot:
        def walk(self):
            raise AssertionError("must not walk the full folder tree to resolve by id")

    monkeypatch.setattr(exchange_client_base, "FolderCollection", FakeFolderCollection)
    backend._account = _fake_account_for_folder_resolution(PoisonRoot())

    result = backend._resolve_folder("AAA=")

    assert result is resolved_folder
    assert len(captured["folders"]) == 1
    assert captured["folders"][0].id == "AAA="


def test_resolve_folder_by_id_containing_slash_is_not_treated_as_path(
    settings, monkeypatch
) -> None:
    backend = EWSExchangeBackend(settings)
    resolved_folder = SimpleNamespace(id="AAA=/BBB==", name="Projects")
    captured: dict = {}

    class FakeFolderCollection:
        def __init__(self, account, folders) -> None:
            captured["folders"] = folders

        def resolve(self):
            yield resolved_folder

    class PoisonRoot:
        def walk(self):
            raise AssertionError("must not walk the full folder tree to resolve by id")

        @property
        def children(self):
            raise AssertionError("must not fall back to path traversal when id lookup succeeds")

    monkeypatch.setattr(exchange_client_base, "FolderCollection", FakeFolderCollection)
    backend._account = _fake_account_for_folder_resolution(PoisonRoot())

    result = backend._resolve_folder("AAA=/BBB==")

    assert result is resolved_folder
    assert captured["folders"][0].id == "AAA=/BBB=="


def test_resolve_folder_falls_back_to_name_lookup_when_id_lookup_is_empty(
    settings, monkeypatch
) -> None:
    backend = EWSExchangeBackend(settings)
    child = SimpleNamespace(name="Projects")

    class FakeFolderCollection:
        def __init__(self, account, folders) -> None:
            pass

        def resolve(self):
            return iter(())

    monkeypatch.setattr(exchange_client_base, "FolderCollection", FakeFolderCollection)
    backend._account = _fake_account_for_folder_resolution(SimpleNamespace(children=[child]))

    result = backend._resolve_folder("Projects")

    assert result is child


def test_resolve_folder_falls_back_to_name_lookup_when_id_lookup_errors(
    settings, monkeypatch
) -> None:
    backend = EWSExchangeBackend(settings)
    child = SimpleNamespace(name="Projects")

    class FakeFolderCollection:
        def __init__(self, account, folders) -> None:
            pass

        def resolve(self):
            raise ErrorInvalidIdMalformed("Id is malformed.")

    monkeypatch.setattr(exchange_client_base, "FolderCollection", FakeFolderCollection)
    backend._account = _fake_account_for_folder_resolution(SimpleNamespace(children=[child]))

    result = backend._resolve_folder("Projects")

    assert result is child


def test_resolve_folder_by_id_propagates_auth_failure_instead_of_swallowing_it(
    settings, monkeypatch
) -> None:
    """Regression guard: an auth/network failure while resolving a folder id must
    surface as AuthFailedError, not be treated the same as 'not an id' and
    silently fall through to a path lookup that can only ever misreport
    not_found."""
    backend = EWSExchangeBackend(settings)

    class FakeFolderCollection:
        def __init__(self, account, folders) -> None:
            pass

        def resolve(self):
            raise UnauthorizedError("access denied")

    class PoisonRoot:
        @property
        def children(self):
            raise AssertionError("must not fall back to path traversal on a real auth failure")

    monkeypatch.setattr(exchange_client_base, "FolderCollection", FakeFolderCollection)
    backend._account = _fake_account_for_folder_resolution(PoisonRoot())

    with pytest.raises(AuthFailedError):
        backend._resolve_folder("AAA=")


def test_resolve_folder_archive_propagates_auth_failure_instead_of_falling_back_to_root(
    settings,
) -> None:
    root = SimpleNamespace()

    class NoArchiveAccount(SimpleNamespace):
        @property
        def archive_root(self):
            raise UnauthorizedError("access denied")

    backend = EWSExchangeBackend(settings)
    backend._account = NoArchiveAccount(
        root=root,
        inbox=object(),
        sent=object(),
        drafts=object(),
        trash=object(),
        junk=object(),
        calendar=object(),
        contacts=object(),
    )

    with pytest.raises(AuthFailedError):
        backend._resolve_folder("archive")


def test_reply_email_sends_response_object_when_no_attachments(settings) -> None:
    backend = EWSExchangeBackend(settings)
    events: list[tuple] = []

    class FakeResponse:
        def send(self):
            events.append(("send",))

    class FakeItem:
        subject = "Hello"

        def create_reply(self, subject, body):
            events.append(("create_reply", subject, body))
            return FakeResponse()

    def fetch(ids, folder=None):
        yield FakeItem()

    backend._account = SimpleNamespace(fetch=fetch)

    result = backend.reply_email(ReplyEmailRequest(id="msg-1", body="Reply body"))

    assert events == [("create_reply", "Re: Hello", "Reply body"), ("send",)]
    assert result == SendResult(id=None, status="sent")


def test_reply_email_saves_draft_attaches_files_then_sends(settings, tmp_path) -> None:
    backend = EWSExchangeBackend(settings)
    attachment_path = tmp_path / "note.txt"
    attachment_path.write_text("hi")
    events: list[tuple] = []
    drafts_folder = SimpleNamespace(id="drafts-folder")

    class FakeMessage:
        id = "sent-1"
        parent_folder_id = SimpleNamespace(id="drafts-folder")

        def attach(self, attachment):
            events.append(("attach", attachment.name))

        def send(self):
            events.append(("send",))

    class FakeDraft:
        id = "draft-1"

    class FakeResponse:
        def save(self, folder):
            events.append(("save", folder))
            return FakeDraft()

    class FakeItem:
        subject = "Hello"

        def create_reply_all(self, subject, body):
            events.append(("create_reply_all", subject, body))
            return FakeResponse()

    fetch_results = iter([FakeItem(), FakeMessage()])

    def fetch(ids, folder=None):
        yield next(fetch_results)

    backend._account = SimpleNamespace(fetch=fetch, drafts=drafts_folder)

    result = backend.reply_email(
        ReplyEmailRequest(
            id="msg-1", body="Reply body", reply_all=True, attachments=[attachment_path]
        )
    )

    assert events == [
        ("create_reply_all", "Re: Hello", "Reply body"),
        ("save", drafts_folder),
        ("attach", "note.txt"),
        ("send",),
    ]
    assert result == SendResult(id=None, status="sent")


@pytest.mark.parametrize(
    ("original", "expected"),
    [
        ("Hello", "Re: Hello"),
        # Replying inside a thread must not stack markers: a few rounds of
        # "Re: Re: Re: " and the subject is unreadable.
        ("Re: Hello", "Re: Hello"),
        ("RE: Hello", "RE: Hello"),
        ("Re[2]: Hello", "Re[2]: Hello"),
        ("Ответ: Привет", "Ответ: Привет"),
        # A forward marker is not a reply marker -- Outlook sends "RE: FW: ...".
        ("FW: Hello", "Re: FW: Hello"),
        ("", "Re:"),
    ],
)
def test_reply_email_does_not_stack_reply_prefixes(settings, original, expected) -> None:
    backend = EWSExchangeBackend(settings)
    captured: list[str] = []

    class FakeResponse:
        def send(self):
            pass

    class FakeItem:
        subject = original

        def create_reply(self, subject, body):
            captured.append(subject)
            return FakeResponse()

    def fetch(ids, folder=None):
        yield FakeItem()

    backend._account = SimpleNamespace(fetch=fetch)

    backend.reply_email(ReplyEmailRequest(id="msg-1", body="Reply body"))

    assert captured == [expected]


@pytest.mark.parametrize(
    ("original", "expected"),
    [
        ("Hello", "Fwd: Hello"),
        ("Fwd: Hello", "Fwd: Hello"),
        ("FW: Hello", "FW: Hello"),
        ("Пересл: Привет", "Пересл: Привет"),
        ("Re: Hello", "Fwd: Re: Hello"),
    ],
)
def test_forward_email_does_not_stack_forward_prefixes(settings, original, expected) -> None:
    backend = EWSExchangeBackend(settings)
    captured: list[str] = []

    class FakeResponse:
        def send(self):
            pass

    class FakeItem:
        subject = original

        def create_forward(self, subject, body, to_recipients):
            captured.append(subject)
            return FakeResponse()

    def fetch(ids, folder=None):
        yield FakeItem()

    backend._account = SimpleNamespace(fetch=fetch)

    backend.forward_email(ForwardEmailRequest(id="msg-1", to=["user@example.com"]))

    assert captured == [expected]


def test_forward_email_saves_draft_attaches_files_then_sends(settings, tmp_path) -> None:
    backend = EWSExchangeBackend(settings)
    attachment_path = tmp_path / "note.txt"
    attachment_path.write_text("hi")
    events: list[tuple] = []
    drafts_folder = SimpleNamespace(id="drafts-folder")

    class FakeMessage:
        id = "sent-1"
        parent_folder_id = SimpleNamespace(id="drafts-folder")

        def attach(self, attachment):
            events.append(("attach", attachment.name))

        def send(self):
            events.append(("send",))

    class FakeDraft:
        id = "draft-1"

    class FakeResponse:
        def save(self, folder):
            events.append(("save", folder))
            return FakeDraft()

    class FakeItem:
        subject = "Hello"

        def create_forward(self, subject, body, to_recipients):
            events.append(
                ("create_forward", subject, body, [m.email_address for m in to_recipients])
            )
            return FakeResponse()

    fetch_results = iter([FakeItem(), FakeMessage()])

    def fetch(ids, folder=None):
        yield next(fetch_results)

    backend._account = SimpleNamespace(fetch=fetch, drafts=drafts_folder)

    result = backend.forward_email(
        ForwardEmailRequest(id="msg-1", to=["dest@example.com"], attachments=[attachment_path])
    )

    assert events == [
        ("create_forward", "Fwd: Hello", "", ["dest@example.com"]),
        ("save", drafts_folder),
        ("attach", "note.txt"),
        ("send",),
    ]
    assert result == SendResult(id=None, status="sent")


@pytest.mark.skipif(not hasattr(os, "mkfifo"), reason="FIFOs are POSIX-only")
def test_attach_files_rejects_fifo_without_hanging(settings, tmp_path) -> None:
    """A FIFO with no writer connected would block open() forever without O_NONBLOCK."""
    backend = EWSExchangeBackend(settings)
    fifo_path = tmp_path / "pipe"
    os.mkfifo(fifo_path)

    class FakeMessage:
        def attach(self, attachment):
            raise AssertionError("should not attach a FIFO")

    outcome: dict = {}

    def run() -> None:
        try:
            backend._attach_files(FakeMessage(), [fifo_path])
        except BaseException as exc:  # noqa: BLE001
            outcome["error"] = exc

    thread = threading.Thread(target=run, daemon=True)
    thread.start()
    thread.join(timeout=5)

    assert not thread.is_alive(), "_attach_files hung opening a FIFO with no writer"
    error = outcome.get("error")
    assert isinstance(error, APIError)
    assert error.code == "validation_error"
    assert "not a regular file" in error.details[0]["reason"]


@pytest.mark.skipif(os.name != "posix", reason="/dev/null device check is POSIX-specific")
def test_attach_files_rejects_character_device(settings) -> None:
    backend = EWSExchangeBackend(settings)

    class FakeMessage:
        def attach(self, attachment):
            raise AssertionError("should not attach a device file")

    with pytest.raises(APIError) as excinfo:
        backend._attach_files(FakeMessage(), [Path(os.devnull)])

    assert excinfo.value.code == "validation_error"
    assert "not a regular file" in excinfo.value.details[0]["reason"]


def test_attach_files_bounds_read_and_reports_oversize_from_actual_bytes(
    settings, tmp_path
) -> None:
    """The size check must come from bytes actually read, not the earlier path-based stat."""
    settings.attachment_max_size_mb = 1
    backend = EWSExchangeBackend(settings)
    oversized = tmp_path / "big.bin"
    oversized.write_bytes(b"x" * (1024 * 1024 + 1))

    class FakeMessage:
        def attach(self, attachment):
            raise AssertionError("oversized attachment must not be attached")

    with pytest.raises(APIError) as excinfo:
        backend._attach_files(FakeMessage(), [oversized])

    assert excinfo.value.code == "validation_error"
    assert "EXCHANGE_ATTACHMENT_MAX_SIZE_MB=1" in excinfo.value.details[0]["reason"]


def test_move_email_returns_new_id_mutated_in_place_by_move(settings) -> None:
    """exchangelib's Item.move() returns None but mutates item.id/changekey in place."""
    backend = EWSExchangeBackend(settings)
    account = _fake_account_for_folder_resolution(object())

    class FakeItem:
        id = "msg-1"

        def move(self, to_folder):
            assert to_folder is account.inbox
            self.id = "msg-1-moved"

    def fetch(ids, folder=None):
        yield FakeItem()

    account.fetch = fetch
    backend._account = account

    result = backend.move_email(FolderActionRequest(id="msg-1", folder="Inbox"))

    assert result == ActionResult(id="msg-1-moved", status="moved", new_folder="Inbox")


def test_copy_email_returns_new_id_from_id_changekey_tuple(settings) -> None:
    """exchangelib's Item.copy() returns an (id, changekey) tuple, not an object with .id."""
    backend = EWSExchangeBackend(settings)
    account = _fake_account_for_folder_resolution(object())

    class FakeItem:
        id = "msg-1"

        def copy(self, to_folder):
            assert to_folder is account.inbox
            return ("msg-1-copy", "changekey-abc")

    def fetch(ids, folder=None):
        yield FakeItem()

    account.fetch = fetch
    backend._account = account

    result = backend.copy_email(FolderActionRequest(id="msg-1", folder="Inbox"))

    assert result == ActionResult(
        id="msg-1", status="copied", new_folder="Inbox", new_id="msg-1-copy"
    )


def test_copy_email_handles_no_result_for_cross_mailbox_copy(settings) -> None:
    """exchangelib returns None from copy() when to_folder is a public/other-mailbox folder."""
    backend = EWSExchangeBackend(settings)
    account = _fake_account_for_folder_resolution(object())

    class FakeItem:
        id = "msg-1"

        def copy(self, to_folder):
            return None

    def fetch(ids, folder=None):
        yield FakeItem()

    account.fetch = fetch
    backend._account = account

    result = backend.copy_email(FolderActionRequest(id="msg-1", folder="Inbox"))

    assert result == ActionResult(id="msg-1", status="copied", new_folder="Inbox", new_id=None)


class _FakeFolder:
    def __init__(self, name: str, id_: str = "folder-1", is_distinguished: bool = False) -> None:
        self.name = name
        self.id = id_
        self.is_distinguished = is_distinguished
        self.parent = None
        self.saved_update_fields: list[str] | None = None
        self.delete_type: str | None = None

    def save(self, update_fields=None):
        self.saved_update_fields = update_fields

    def delete(self, delete_type=None):
        self.delete_type = delete_type


def test_rename_folder_renames_and_returns_new_path(settings) -> None:
    backend = EWSExchangeBackend(settings)
    folder = _FakeFolder(name="Марафон")
    account = _fake_account_for_folder_resolution(object())
    account.inbox = folder
    backend._account = account

    result = backend.rename_folder(RenameFolderRequest(folder="inbox", name="Marathon"))

    assert folder.name == "Marathon"
    assert folder.saved_update_fields == ["name"]
    assert result == ActionResult(id="folder-1", status="renamed", path="Marathon")


def test_rename_folder_rejects_distinguished_folder(settings) -> None:
    """Renaming Inbox/Calendar/etc. would succeed server-side but is never what a
    caller actually wants -- refuse it client-side with a clear error instead."""
    backend = EWSExchangeBackend(settings)
    folder = _FakeFolder(name="Inbox", is_distinguished=True)
    account = _fake_account_for_folder_resolution(object())
    account.inbox = folder
    backend._account = account

    with pytest.raises(APIError) as excinfo:
        backend.rename_folder(RenameFolderRequest(folder="inbox", name="Something"))

    assert excinfo.value.code == "validation_error"
    assert folder.name == "Inbox"
    assert folder.saved_update_fields is None


def test_rename_folder_maps_exception_from_save(settings) -> None:
    backend = EWSExchangeBackend(settings)

    class FailingFolder(_FakeFolder):
        def save(self, update_fields=None):
            raise UnauthorizedError("access denied")

    folder = FailingFolder(name="Old")
    account = _fake_account_for_folder_resolution(object())
    account.inbox = folder
    backend._account = account

    with pytest.raises(AuthFailedError):
        backend.rename_folder(RenameFolderRequest(folder="inbox", name="New"))


def test_delete_folder_soft_deletes_by_default(settings) -> None:
    backend = EWSExchangeBackend(settings)
    folder = _FakeFolder(name="Old Folder", id_="folder-99")
    account = _fake_account_for_folder_resolution(object())
    account.inbox = folder
    backend._account = account

    result = backend.delete_folder(DeleteFolderRequest(folder="inbox"))

    assert folder.delete_type == MOVE_TO_DELETED_ITEMS
    assert result == ActionResult(id="folder-99", status="deleted")


def test_delete_folder_hard_deletes_when_requested(settings) -> None:
    backend = EWSExchangeBackend(settings)
    folder = _FakeFolder(name="Old Folder", id_="folder-99")
    account = _fake_account_for_folder_resolution(object())
    account.inbox = folder
    backend._account = account

    backend.delete_folder(DeleteFolderRequest(folder="inbox", hard_delete=True))

    assert folder.delete_type == HARD_DELETE


def test_delete_folder_rejects_distinguished_folder(settings) -> None:
    backend = EWSExchangeBackend(settings)
    folder = _FakeFolder(name="Inbox", is_distinguished=True)
    account = _fake_account_for_folder_resolution(object())
    account.inbox = folder
    backend._account = account

    with pytest.raises(APIError) as excinfo:
        backend.delete_folder(DeleteFolderRequest(folder="inbox"))

    assert excinfo.value.code == "validation_error"
    assert folder.delete_type is None


def test_delete_folder_maps_exception_from_delete(settings) -> None:
    backend = EWSExchangeBackend(settings)

    class FailingFolder(_FakeFolder):
        def delete(self, delete_type=None):
            raise UnauthorizedError("access denied")

    folder = FailingFolder(name="Old")
    account = _fake_account_for_folder_resolution(object())
    account.inbox = folder
    backend._account = account

    with pytest.raises(AuthFailedError):
        backend.delete_folder(DeleteFolderRequest(folder="inbox"))


def _bound_inbox() -> Inbox:
    """A real Inbox folder bound to just enough fake account to compile EWS restriction XML."""
    account = SimpleNamespace(version=Version(build=EXCHANGE_2016))
    root = SimpleNamespace(account=account, is_deleteable=False)
    return Inbox(root=root)


def test_list_emails_from_address_filter_uses_author_field_not_a_subfield(settings) -> None:
    """'author' is a MailboxField, not an IndexedField: EWS rejects a '__email_address' subfield path
    with 'Unknown field path', so filtering must target 'author' itself (see the __iexact lookup)."""
    backend = EWSExchangeBackend(settings)
    captured: dict = {}

    class FakeQuerySet:
        def only(self, *fields):
            return self

        def order_by(self, *args):
            return self

        def filter(self, **filters):
            captured.update(filters)
            return self

        def __getitem__(self, item):
            return []

    account = _fake_account_for_folder_resolution(object())
    account.inbox = SimpleNamespace(all=lambda: FakeQuerySet())
    backend._account = account

    backend.list_emails(ListEmailsRequest(folder="Inbox", from_address="foo@example.com"))

    assert "author__email_address" not in captured
    assert captured["author__iexact"] == "foo@example.com"

    inbox = _bound_inbox()
    restriction = Restriction(Q(**captured), folders=[inbox], applies_to=Restriction.ITEMS)
    xml = restriction.to_xml(version=inbox.account.version)
    assert (
        xml.find(".//{http://schemas.microsoft.com/exchange/services/2006/types}FieldURI").get(
            "FieldURI"
        )
        == "message:From"
    )


def test_author_email_address_subfield_path_is_rejected_by_ews() -> None:
    """Regression guard: confirms the old, buggy filter key really is invalid EWS syntax."""
    inbox = _bound_inbox()
    restriction = Restriction(
        Q(author__email_address="foo@example.com"), folders=[inbox], applies_to=Restriction.ITEMS
    )

    with pytest.raises(Exception, match="Unknown field path 'author__email_address'"):
        restriction.to_xml(version=inbox.account.version)


def test_search_emails_raises_auth_failed_instead_of_swallowing_it(settings) -> None:
    """Regression guard: an auth failure while searching must surface as AuthFailedError,
    not get treated as 'no results'."""
    backend = EWSExchangeBackend(settings)

    class FailingQuerySet:
        def order_by(self, *args):
            return self

        def filter(self, *args, **kwargs):
            raise UnauthorizedError("access denied")

    account = _fake_account_for_folder_resolution(object())
    account.inbox = FailingQuerySet()
    backend._account = account

    with pytest.raises(AuthFailedError):
        backend.search_emails(SearchEmailsRequest(query="hello", folder="inbox"))


def test_search_emails_matches_subject_or_body_or_sender_in_one_pass(settings) -> None:
    """Regression guard: search_emails used to check subject first and only fall
    back to text_body when subject had zero hits, silently dropping body-only
    matches whenever anything else in the folder matched on subject. The
    filter must now search subject, text_body, and sender together."""
    backend = EWSExchangeBackend(settings)
    captured: dict = {}

    class FakeQuerySet:
        def filter(self, restriction):
            captured["restriction"] = restriction
            return self

        def only(self, *fields):
            return self

        def order_by(self, *args):
            return self

        def __getitem__(self, item):
            return []

    account = _fake_account_for_folder_resolution(object())
    account.inbox = FakeQuerySet()
    backend._account = account

    backend.search_emails(SearchEmailsRequest(query="hello", folder="inbox"))

    restriction_text = str(captured["restriction"])
    assert "subject" in restriction_text
    assert "text_body" in restriction_text
    assert "author" in restriction_text


def test_bulk_move_reports_per_item_results_without_failing_the_batch(settings) -> None:
    """One bad id out of a batch must not undo or hide the moves that worked."""
    backend = EWSExchangeBackend(settings)

    def bulk_move(ids, to_folder):
        return [("new-1", "ck"), ErrorItemNotFound("gone"), ("new-3", "ck")]

    account = _fake_account_for_folder_resolution(object())
    account.inbox = object()
    account.bulk_move = bulk_move
    backend._account = account

    from outlook_mcp.models import BulkMoveEmailsRequest

    result = backend.move_emails(BulkMoveEmailsRequest(ids=["a", "b", "c"], folder="inbox"))

    assert [(r.id, r.new_id) for r in result.succeeded] == [("a", "new-1"), ("c", "new-3")]
    assert [(f.id, f.error) for f in result.failed] == [("b", "not_found")]


def test_bulk_delete_maps_soft_and_hard_delete_types(settings) -> None:
    backend = EWSExchangeBackend(settings)
    captured: list[str] = []

    def bulk_delete(ids, delete_type):
        captured.append(delete_type)
        return [True for _ in ids]

    backend._account = SimpleNamespace(bulk_delete=bulk_delete)

    from outlook_mcp.models import BulkDeleteEmailsRequest

    backend.delete_emails(BulkDeleteEmailsRequest(ids=["a"]))
    backend.delete_emails(BulkDeleteEmailsRequest(ids=["a"], hard_delete=True))

    assert captured == [MOVE_TO_DELETED_ITEMS, HARD_DELETE]


def test_bulk_move_to_a_public_folder_still_gives_every_id_a_verdict(settings) -> None:
    """Moving into a public folder returns no per-item ids at all; every input
    id must still come back as succeeded rather than the call crashing."""
    backend = EWSExchangeBackend(settings)
    account = _fake_account_for_folder_resolution(object())
    account.inbox = object()
    account.bulk_move = lambda ids, to_folder: []
    backend._account = account

    from outlook_mcp.models import BulkMoveEmailsRequest

    result = backend.move_emails(BulkMoveEmailsRequest(ids=["a", "b"], folder="inbox"))

    assert [r.id for r in result.succeeded] == ["a", "b"]
    assert result.failed == []


def test_search_emails_aqs_sends_a_query_string_not_a_restriction(settings) -> None:
    """AQS goes to EWS as its own FindItem QueryString element and cannot be
    combined with a Restriction, so the substring filters must not be added."""
    backend = EWSExchangeBackend(settings)
    captured: dict = {}

    class FakeQuerySet:
        def filter(self, restriction):
            captured["restriction"] = restriction
            return self

        def only(self, *fields):
            return self

        def order_by(self, *args):
            return self

        def __getitem__(self, item):
            return []

    account = _fake_account_for_folder_resolution(object())
    account.inbox = FakeQuerySet()
    backend._account = account

    backend.search_emails(
        SearchEmailsRequest(aqs="from:ivan AND hasattachments:true", folder="inbox")
    )

    q = captured["restriction"]
    assert q.query_string == "from:ivan AND hasattachments:true"
    assert "icontains" not in str(q)


def test_search_emails_requires_exactly_one_of_query_or_aqs() -> None:
    with pytest.raises(ValueError, match="exactly one of query or aqs"):
        SearchEmailsRequest(query="hello", aqs="from:ivan")

    with pytest.raises(ValueError, match="exactly one of query or aqs"):
        SearchEmailsRequest(folder="inbox")


def test_search_emails_defaults_to_mail_folders_only(settings, monkeypatch) -> None:
    """Regression guard: search_emails used to default to Inbox only. It should
    now search the whole mailbox by default, but only actual mail folders
    (folder_class IPF.Note), not e.g. Calendar or Contacts."""
    backend = EWSExchangeBackend(settings)
    mail_folder = SimpleNamespace(folder_class="IPF.Note", name="Inbox")
    other_folder = SimpleNamespace(folder_class="IPF.Appointment", name="Calendar")
    captured: dict = {}

    class FakeRoot:
        def walk(self):
            return [mail_folder, other_folder]

    class FakeFolderCollection:
        def __init__(self, account, folders) -> None:
            captured["folders"] = folders

        def filter(self, restriction):
            return self

        def only(self, *fields):
            return self

        def order_by(self, *args):
            return self

        def __getitem__(self, item):
            return []

    monkeypatch.setattr(exchange_client_email, "FolderCollection", FakeFolderCollection)
    backend._account = SimpleNamespace(root=FakeRoot())

    backend.search_emails(SearchEmailsRequest(query="hello"))

    assert captured["folders"] == [mail_folder]


def test_search_emails_maps_error_raised_while_walking_folders(settings) -> None:
    """Regression guard: account.root.walk() used to run outside the try/except,
    so an auth/network failure while walking the mailbox for a whole-mailbox
    search escaped as a raw exchangelib exception instead of an APIError."""
    backend = EWSExchangeBackend(settings)

    class PoisonRoot:
        def walk(self):
            raise UnauthorizedError("access denied")

    backend._account = SimpleNamespace(root=PoisonRoot())

    with pytest.raises(AuthFailedError):
        backend.search_emails(SearchEmailsRequest(query="hello"))


def test_list_folders_maps_error_raised_while_walking_children(settings) -> None:
    """Regression guard: the recursive folder walk in _to_folder_info ran outside
    any try/except, so a failure partway through (e.g. auth expiring mid-walk)
    escaped as a raw exchangelib exception instead of an APIError."""
    from outlook_mcp.models import ListFoldersRequest

    backend = EWSExchangeBackend(settings)

    class PoisonChildren:
        @property
        def children(self):
            raise UnauthorizedError("access denied")

    backend._account = SimpleNamespace(root=PoisonChildren())

    with pytest.raises(AuthFailedError):
        backend.list_folders(ListFoldersRequest(parent="root", depth=2))
