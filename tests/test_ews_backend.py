from __future__ import annotations

from types import SimpleNamespace

import pytest
from exchangelib.errors import ErrorInvalidIdMalformed
from exchangelib.ewsdatetime import EWSTimeZone

import outlook_mcp.exchange_client as exchange_client_module
from outlook_mcp.errors import APIError
from outlook_mcp.exchange_client import EWSExchangeBackend
from outlook_mcp.models import FindFreeSlotsRequest, GetContactRequest, GetEmailRequest


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

    monkeypatch.setattr(exchange_client_module, "ResolveNames", FakeResolveNames)
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

    monkeypatch.setattr(exchange_client_module, "ResolveNames", FakeResolveNames)
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

    monkeypatch.setattr(exchange_client_module, "ResolveNames", FakeResolveNames)
    backend._account = SimpleNamespace(contacts=object(), protocol=object())

    with pytest.raises(APIError) as excinfo:
        backend.get_contact(GetContactRequest(id="nobody@example.com"))

    assert excinfo.value.code == "not_found"
