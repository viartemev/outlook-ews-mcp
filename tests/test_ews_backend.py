from __future__ import annotations

from types import SimpleNamespace

import pytest
from exchangelib.errors import ErrorInvalidIdMalformed
from exchangelib.ewsdatetime import EWSTimeZone
from exchangelib.folders import Inbox
from exchangelib.restriction import Q, Restriction
from exchangelib.version import EXCHANGE_2016, Version

import outlook_mcp.exchange_client.base as exchange_client_base
import outlook_mcp.exchange_client.contacts as exchange_client_contacts
from outlook_mcp.errors import APIError
from outlook_mcp.exchange_client import EWSExchangeBackend
from outlook_mcp.models import (
    ActionResult,
    FindFreeSlotsRequest,
    FolderActionRequest,
    ForwardEmailRequest,
    GetContactRequest,
    GetEmailRequest,
    ListEmailsRequest,
    ReplyEmailRequest,
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
    backend._account = _fake_account(
        [SimpleNamespace(merged="00"), SimpleNamespace(merged="02")]
    )

    slots = backend.find_free_slots(request)

    assert len(slots) == 1
    assert slots[0].start.hour == 9


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


def test_get_contact_resolves_gal_address(settings, monkeypatch) -> None:
    backend = EWSExchangeBackend(settings)
    captured: dict = {}
    gal_contact = SimpleNamespace(
        id=None,
        display_name="Ivan Ivanov",
        file_as=None,
        given_name="Ivan",
        surname="Ivanov",
        email_addresses=[SimpleNamespace(label="EmailAddress1", email="ivan@example.com")],
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
            SimpleNamespace(label="EmailAddress1", email="X500:/o=TANDER/ou=Exchange Administrative Group/cn=ivan"),
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

    assert [entry.address for entry in contact.email_addresses] == ["ivan@example.com", "ivan.alias@example.com"]
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


def test_resolve_folder_by_id_containing_slash_is_not_treated_as_path(settings, monkeypatch) -> None:
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


def test_resolve_folder_falls_back_to_name_lookup_when_id_lookup_is_empty(settings, monkeypatch) -> None:
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


def test_resolve_folder_falls_back_to_name_lookup_when_id_lookup_errors(settings, monkeypatch) -> None:
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
    assert result == SendResult(id="msg-1", status="sent")


def test_reply_email_saves_draft_attaches_files_then_sends(settings, tmp_path) -> None:
    backend = EWSExchangeBackend(settings)
    attachment_path = tmp_path / "note.txt"
    attachment_path.write_text("hi")
    events: list[tuple] = []
    drafts_folder = object()

    class FakeMessage:
        id = "sent-1"

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
    assert result == SendResult(id="sent-1", status="sent")


def test_forward_email_saves_draft_attaches_files_then_sends(settings, tmp_path) -> None:
    backend = EWSExchangeBackend(settings)
    attachment_path = tmp_path / "note.txt"
    attachment_path.write_text("hi")
    events: list[tuple] = []
    drafts_folder = object()

    class FakeMessage:
        id = "sent-1"

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
    assert result == SendResult(id="sent-1", status="sent")


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

    assert result == ActionResult(id="msg-1", status="copied", new_folder="Inbox", new_id="msg-1-copy")


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
    assert xml.find(".//{http://schemas.microsoft.com/exchange/services/2006/types}FieldURI").get(
        "FieldURI"
    ) == "message:From"


def test_author_email_address_subfield_path_is_rejected_by_ews() -> None:
    """Regression guard: confirms the old, buggy filter key really is invalid EWS syntax."""
    inbox = _bound_inbox()
    restriction = Restriction(Q(author__email_address="foo@example.com"), folders=[inbox], applies_to=Restriction.ITEMS)

    with pytest.raises(Exception, match="Unknown field path 'author__email_address'"):
        restriction.to_xml(version=inbox.account.version)
