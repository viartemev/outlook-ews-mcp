from __future__ import annotations

from datetime import UTC, datetime
from types import SimpleNamespace

import pytest

from outlook_mcp.errors import APIError
from outlook_mcp.exchange_client import EWSExchangeBackend
from outlook_mcp.models import (
    MarkEmailRequest,
    UpdateContactRequest,
    UpdateEventRequest,
)


#: Shared folder id used by FakeSavableItem.parent_folder_id and _fake_folder(),
#: so _fetch_item's real-vs-fake-folder scoping check (item.parent_folder_id.id ==
#: folder.id) matches, the way a real fetch against that real folder would.
_FAKE_FOLDER_ID = "folder-1"


def _fake_folder() -> SimpleNamespace:
    return SimpleNamespace(id=_FAKE_FOLDER_ID)


class FakeSavableItem(SimpleNamespace):
    def __init__(self, **kwargs) -> None:
        kwargs.setdefault("parent_folder_id", SimpleNamespace(id=_FAKE_FOLDER_ID))
        super().__init__(**kwargs)
        self.save_calls: list[dict] = []

    def save(self, **kwargs) -> None:
        self.save_calls.append(kwargs)


def _account_with_item(item: FakeSavableItem, **extra) -> SimpleNamespace:
    return SimpleNamespace(fetch=lambda **kwargs: iter([item]), **extra)


def test_mark_email_saves_real_ews_field_names_and_preserves_categories(settings) -> None:
    backend = EWSExchangeBackend(settings)
    item = FakeSavableItem(
        id="email-1", is_read=False, importance="Normal", categories=["Existing"]
    )
    backend._account = _account_with_item(item)

    request = MarkEmailRequest.model_validate({"id": "email-1", "read": True, "flag": "flagged"})
    result = backend.mark_email(request)

    assert item.is_read is True
    assert item.flag_status == 2
    assert item.categories == ["Existing"]  # untouched, not overwritten by flag
    assert item.save_calls[-1]["update_fields"] == ["is_read", "flag_status"]
    assert result.updated_fields == ["read", "flag"]


def test_mark_email_empty_request_does_not_save(settings) -> None:
    """An id-only MarkEmailRequest has nothing to change -- saving anyway would
    still write every loaded field back via item.save(update_fields=None)."""
    backend = EWSExchangeBackend(settings)
    item = FakeSavableItem(id="email-1", is_read=False)
    backend._account = _account_with_item(item)

    request = MarkEmailRequest.model_validate({"id": "email-1"})
    result = backend.mark_email(request)

    assert item.save_calls == []
    assert result.updated_fields == []


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
    calendar = _fake_folder()
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
    assert item.save_calls[-1]["update_fields"] == [
        "reminder_minutes_before_start",
        "required_attendees",
    ]
    assert result.updated_fields == ["reminder_minutes", "add_attendees", "remove_attendees"]


def test_update_event_rejects_lone_start_past_existing_end(settings) -> None:
    backend = EWSExchangeBackend(settings)
    item = FakeSavableItem(
        id="event-1",
        start=datetime(2026, 4, 8, 9, 0, tzinfo=UTC),
        end=datetime(2026, 4, 8, 10, 0, tzinfo=UTC),
    )
    backend._account = _account_with_item(item, calendar=_fake_folder(), default_timezone=UTC)

    request = UpdateEventRequest.model_validate(
        {"id": "event-1", "start": "2026-04-08T11:00:00+00:00"}
    )

    with pytest.raises(APIError) as excinfo:
        backend.update_event(request)
    assert excinfo.value.code == "validation_error"
    assert item.save_calls == []


def test_update_event_empty_update_does_not_save_or_notify(settings) -> None:
    backend = EWSExchangeBackend(settings)
    item = FakeSavableItem(
        id="event-1",
        start=datetime(2026, 4, 8, 9, 0, tzinfo=UTC),
        end=datetime(2026, 4, 8, 10, 0, tzinfo=UTC),
    )
    backend._account = _account_with_item(item, calendar=_fake_folder(), default_timezone=UTC)

    request = UpdateEventRequest.model_validate({"id": "event-1"})
    result = backend.update_event(request)

    assert item.save_calls == []
    assert result.updated_fields == []


def test_update_contact_omitted_fields_left_untouched(settings) -> None:
    backend = EWSExchangeBackend(settings)
    item = FakeSavableItem(id="contact-1", display_name="Old Name", job_title="Old Title")
    backend._account = _account_with_item(item, contacts=_fake_folder())

    request = UpdateContactRequest.model_validate({"id": "contact-1", "job_title": "New Title"})
    backend.update_contact(request)

    assert item.display_name == "Old Name"
    assert item.job_title == "New Title"
    assert item.save_calls[-1]["update_fields"] == ["job_title"]


def test_update_contact_explicit_null_clears_phone(settings) -> None:
    backend = EWSExchangeBackend(settings)
    item = FakeSavableItem(id="contact-1", phone_numbers=[SimpleNamespace(phone_number="+1000")])
    backend._account = _account_with_item(item, contacts=_fake_folder())

    request = UpdateContactRequest.model_validate({"id": "contact-1", "phone": None})
    backend.update_contact(request)

    assert item.phone_numbers == []
    assert item.save_calls[-1]["update_fields"] == ["phone_numbers"]


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
    backend._account = _account_with_item(item, contacts=_fake_folder())

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
