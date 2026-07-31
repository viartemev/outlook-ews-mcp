from __future__ import annotations

from types import SimpleNamespace

from outlook_mcp.exchange_client import EWSExchangeBackend
from outlook_mcp.models import (
    MarkEmailRequest,
    UpdateContactRequest,
    UpdateEventRequest,
)


class FakeSavableItem(SimpleNamespace):
    def __init__(self, **kwargs) -> None:
        super().__init__(**kwargs)
        self.save_calls: list[dict] = []

    def save(self, **kwargs) -> None:
        self.save_calls.append(kwargs)


def _account_with_item(item: FakeSavableItem, **extra) -> SimpleNamespace:
    return SimpleNamespace(fetch=lambda **kwargs: iter([item]), **extra)


def test_mark_email_saves_real_ews_field_names_and_preserves_categories(settings) -> None:
    backend = EWSExchangeBackend(settings)
    item = FakeSavableItem(id="email-1", is_read=False, importance="Normal", categories=["Existing"])
    backend._account = _account_with_item(item)

    request = MarkEmailRequest.model_validate({"id": "email-1", "read": True, "flag": "flagged"})
    result = backend.mark_email(request)

    assert item.is_read is True
    assert item.flag_status == 2
    assert item.categories == ["Existing"]  # untouched, not overwritten by flag
    assert item.save_calls[-1]["update_fields"] == ["is_read", "flag_status"]
    assert result.updated_fields == ["read", "flag"]


def test_mark_email_flag_none_clears_status(settings) -> None:
    backend = EWSExchangeBackend(settings)
    item = FakeSavableItem(id="email-1", flag_status=2)
    backend._account = _account_with_item(item)

    request = MarkEmailRequest.model_validate({"id": "email-1", "flag": "none"})
    backend.mark_email(request)

    assert item.flag_status is None
    assert item.save_calls[-1]["update_fields"] == ["flag_status"]


def test_update_event_saves_real_ews_field_names(settings) -> None:
    backend = EWSExchangeBackend(settings)
    item = FakeSavableItem(id="event-1", reminder_minutes_before_start=15, required_attendees=[])
    calendar = object()
    backend._account = _account_with_item(item, calendar=calendar)

    request = UpdateEventRequest.model_validate(
        {
            "id": "event-1",
            "reminder_minutes": 30,
            "add_attendees": ["new@example.com"],
            "remove_attendees": ["old@example.com"],
        }
    )
    result = backend.update_event(request)

    assert item.reminder_minutes_before_start == 30
    assert item.save_calls[-1]["update_fields"] == ["reminder_minutes_before_start", "required_attendees"]
    assert result.updated_fields == ["reminder_minutes", "add_attendees", "remove_attendees"]


def test_update_contact_saves_real_ews_field_names(settings) -> None:
    backend = EWSExchangeBackend(settings)
    item = FakeSavableItem(
        id="contact-1",
        display_name=None,
        given_name=None,
        company_name=None,
        email_addresses=[],
        phone_numbers=[],
    )
    backend._account = _account_with_item(item, contacts=object())

    request = UpdateContactRequest.model_validate(
        {
            "id": "contact-1",
            "display_name": "Ivan Petrov",
            "first_name": "Ivan",
            "company": "Acme",
            "email": "ivan@example.com",
            "phone": "+1000",
        }
    )
    result = backend.update_contact(request)

    assert item.given_name == "Ivan"
    assert item.company_name == "Acme"
    assert item.save_calls[-1]["update_fields"] == [
        "display_name",
        "given_name",
        "company_name",
        "email_addresses",
        "phone_numbers",
    ]
    assert result.updated_fields == ["display_name", "first_name", "company", "email", "phone"]
