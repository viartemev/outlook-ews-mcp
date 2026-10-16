"""Serialize what our backend actually builds through real exchangelib objects.

Unlike the SimpleNamespace-based fakes elsewhere, these tests exercise the real
exchangelib field descriptors so that setting a read-only or unsupported EWS
field is caught: exchangelib accepts such values silently at construction time
and only drops them when it renders the create/update XML (see
``Item.to_xml``), so a fake object that just stores whatever kwargs it was
given can't catch that class of bug.

``backend._account`` is a real (but never-initialized) ``exchangelib.Account``
instance -- built via ``Account.__new__`` to satisfy exchangelib's own
``isinstance`` checks on the ``account=``/``folder=`` constructor kwargs
without performing autodiscover or any network I/O. Folder attributes are left
``None`` so the constructors receive ``folder=None``, which exchangelib also
accepts without touching a live folder tree.
"""

from __future__ import annotations

from datetime import UTC

from exchangelib import Account, CalendarItem, Contact, Message
from exchangelib.version import EXCHANGE_2016, Version
from lxml import etree

from outlook_mcp.exchange_client import EWSExchangeBackend
from outlook_mcp.models import (
    CreateContactRequest,
    CreateEventRequest,
    SendEmailRequest,
    UpdateContactRequest,
)

VERSION = Version(build=EXCHANGE_2016)


def _to_xml(item) -> str:
    return etree.tostring(item.to_xml(version=VERSION)).decode()


def _bare_account(**folders) -> Account:
    account = Account.__new__(Account)
    account.default_timezone = UTC
    for name in ("calendar", "drafts", "contacts"):
        setattr(account, name, folders.get(name))
    return account


def _capture_save(monkeypatch, cls) -> dict:
    captured: dict = {}

    def fake_save(self, **kwargs):
        captured["xml"] = _to_xml(self)
        self.id = "new-id"
        return self.id

    monkeypatch.setattr(cls, "save", fake_save)
    return captured


def test_create_event_omits_read_only_fields_from_xml(settings, monkeypatch) -> None:
    """Regression test: ``is_online_meeting`` is server-computed and read-only
    in exchangelib, so a value passed for it is silently dropped from the
    create payload rather than raising -- ``CreateEventRequest`` no longer
    exposes it (see exchange_client/calendar.py create_event)."""
    captured = _capture_save(monkeypatch, CalendarItem)
    backend = EWSExchangeBackend(settings)
    backend._account = _bare_account()

    request = CreateEventRequest.model_validate(
        {
            "subject": "Sprint planning",
            "start": "2026-04-13T09:00:00+00:00",
            "end": "2026-04-13T10:00:00+00:00",
            "attendees": ["ivan@example.com"],
            "importance": "high",
        }
    )

    backend.create_event(request)

    xml = captured["xml"]
    assert "Sprint planning" in xml
    assert "<t:Importance>High</t:Importance>" in xml
    assert "ivan@example.com" in xml
    # Read-only/server-computed fields must never appear in outbound XML.
    for read_only_field in ("IsOnlineMeeting", "Organizer", "IsRecurring"):
        assert read_only_field not in xml


def test_create_event_with_meeting_url_sets_net_show_url_and_body(settings, monkeypatch) -> None:
    captured = _capture_save(monkeypatch, CalendarItem)
    backend = EWSExchangeBackend(settings)
    backend._account = _bare_account()

    request = CreateEventRequest.model_validate(
        {
            "subject": "Standup",
            "start": "2026-04-13T09:00:00+00:00",
            "end": "2026-04-13T09:15:00+00:00",
            "body": "Agenda: status updates",
            "meeting_url": "https://teams.microsoft.com/l/meetup-join/abc",
        }
    )

    backend.create_event(request)

    xml = captured["xml"]
    assert "<t:NetShowUrl>https://teams.microsoft.com/l/meetup-join/abc</t:NetShowUrl>" in xml
    assert "Join the meeting: https://teams.microsoft.com/l/meetup-join/abc" in xml
    assert "Agenda: status updates" in xml
    assert "<t:Location>https://teams.microsoft.com/l/meetup-join/abc</t:Location>" in xml


def test_create_event_with_meeting_url_and_room_combines_location(settings, monkeypatch) -> None:
    """A real room stays visible after the link, matching how the mailbox's own
    online-meeting add-ins compose location (link first, room kept after)."""
    captured = _capture_save(monkeypatch, CalendarItem)
    backend = EWSExchangeBackend(settings)
    backend._account = _bare_account()

    request = CreateEventRequest.model_validate(
        {
            "subject": "Standup",
            "start": "2026-04-13T09:00:00+00:00",
            "end": "2026-04-13T09:15:00+00:00",
            "location": "520-Москва Tower A-Переговорная",
            "meeting_url": "https://talk.magnit.ru/abc123",
        }
    )

    backend.create_event(request)

    # Cyrillic text comes back as numeric XML entities in the raw string, so
    # compare the parsed element text rather than a literal substring.
    tree = etree.fromstring(captured["xml"].encode())
    location = tree.find(".//{http://schemas.microsoft.com/exchange/services/2006/types}Location")
    assert location.text == "https://talk.magnit.ru/abc123; 520-Москва Tower A-Переговорная"


def test_send_email_serializes_recipients_and_importance(settings, monkeypatch) -> None:
    captured = _capture_save(monkeypatch, Message)
    backend = EWSExchangeBackend(settings)
    backend._account = _bare_account()

    request = SendEmailRequest.model_validate(
        {
            "to": ["user@example.com"],
            "cc": ["cc@example.com"],
            "subject": "Hello",
            "body": "World",
            "importance": "low",
        }
    )

    message = backend._make_message(request)
    message.save()

    xml = captured["xml"]
    assert "<t:Subject>Hello</t:Subject>" in xml
    assert "user@example.com" in xml
    assert "cc@example.com" in xml
    assert "<t:Importance>Low</t:Importance>" in xml


def test_create_contact_serializes_indexed_properties(settings, monkeypatch) -> None:
    captured = _capture_save(monkeypatch, Contact)
    backend = EWSExchangeBackend(settings)
    backend._account = _bare_account()

    request = CreateContactRequest.model_validate(
        {
            "display_name": "Ivan Ivanov",
            "email": "ivan@example.com",
            "phone": "+79990000000",
        }
    )

    backend.create_contact(request)

    xml = captured["xml"]
    assert "Ivan Ivanov" in xml
    assert "ivan@example.com" in xml
    assert "+79990000000" in xml


def test_create_contact_notes_serialize_via_body_not_readonly_notes_field(
    settings, monkeypatch
) -> None:
    """Regression test: ``Contact.notes`` is read-only in exchangelib and would
    be silently dropped from the create payload -- Outlook actually stores
    contact notes in the item body, so ``create_contact`` writes there
    instead (see exchange_client/contacts.py)."""
    captured = _capture_save(monkeypatch, Contact)
    backend = EWSExchangeBackend(settings)
    backend._account = _bare_account()

    request = CreateContactRequest.model_validate(
        {"display_name": "Ivan Ivanov", "notes": "Met at the conference"}
    )

    backend.create_contact(request)

    xml = captured["xml"]
    assert "Met at the conference" in xml
    assert "<t:Notes>" not in xml


def test_update_contact_only_serializes_requested_fields(settings, monkeypatch) -> None:
    contact = Contact(display_name="Old Name")
    backend = EWSExchangeBackend(settings)
    backend._account = _bare_account()
    backend._account.fetch = lambda **kwargs: iter([contact])
    monkeypatch.setattr(Contact, "save", lambda self, **kwargs: None)

    request = UpdateContactRequest.model_validate(
        {"id": "contact-1", "display_name": "Old Name", "job_title": "Manager"}
    )
    backend.update_contact(request)

    xml = _to_xml(contact)
    assert "Old Name" in xml
    assert "Manager" in xml
