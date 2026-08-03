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
