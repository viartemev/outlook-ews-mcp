from __future__ import annotations

from datetime import UTC, datetime
from types import SimpleNamespace

import pytest
from exchangelib import DELEGATE
from exchangelib.errors import ErrorAccessDenied
from pydantic import ValidationError

import outlook_mcp.exchange_client.base as exchange_client_base
import outlook_mcp.exchange_client.calendar as exchange_client_calendar
from outlook_mcp.config import Settings
from outlook_mcp.errors import APIError, PermissionDeniedError
from outlook_mcp.exchange_client import EWSExchangeBackend
from outlook_mcp.models import GetEventRequest, ListEventsRequest


@pytest.fixture
def settings() -> Settings:
    # An "@"-containing username so build_auth_context (called from
    # _account_for) can determine a primary SMTP address without needing
    # EXCHANGE_EMAIL_ADDRESS/EXCHANGE_IMPERSONATE_AS set.
    return Settings(
        _env_file=None,
        EXCHANGE_SERVER="https://mail.example.com/EWS/Exchange.asmx",
        EXCHANGE_USERNAME="user@example.com",
        EXCHANGE_PASSWORD="secret",
    )


class FakeAccount:
    instances: list[FakeAccount] = []

    def __init__(self, primary_smtp_address, config, autodiscover, access_type):
        self.primary_smtp_address = primary_smtp_address
        self.config = config
        self.autodiscover = autodiscover
        self.access_type = access_type
        self.protocol = SimpleNamespace(raw_session=lambda *a, **k: None)
        FakeAccount.instances.append(self)


def test_account_for_builds_and_caches_per_mailbox(settings, monkeypatch) -> None:
    FakeAccount.instances = []
    monkeypatch.setattr(exchange_client_base, "Account", FakeAccount)
    backend = EWSExchangeBackend(settings)
    fake_config = object()
    backend._account = SimpleNamespace(protocol=SimpleNamespace(config=fake_config))

    account_a = backend._account_for("colleague@example.com")
    account_b = backend._account_for("colleague@example.com")
    backend._account_for("other@example.com")

    assert account_a is account_b
    assert len(FakeAccount.instances) == 2
    first = FakeAccount.instances[0]
    assert first.primary_smtp_address == "colleague@example.com"
    assert first.config is fake_config
    assert first.autodiscover is False
    assert first.access_type == DELEGATE


def test_account_for_maps_exceptions_from_account_construction(settings, monkeypatch) -> None:
    class FailingAccount:
        def __init__(self, *args, **kwargs) -> None:
            raise RuntimeError("boom")

    monkeypatch.setattr(exchange_client_base, "Account", FailingAccount)
    backend = EWSExchangeBackend(settings)
    backend._account = SimpleNamespace(protocol=SimpleNamespace(config=object()))

    with pytest.raises(APIError):
        backend._account_for("colleague@example.com")


class _CalendarOnlyAccount:
    """A mailbox shared with calendar rights only: Account.calendar would batch-resolve
    every distinguished folder and fail, so touching it is a bug."""

    @property
    def calendar(self):
        raise AssertionError("must not resolve the mailbox's distinguished folders")


def _fake_single_folder_query_set(result, captured):
    class FakeSingleFolderQuerySet:
        def __init__(self, account, folder):
            captured["account"] = account
            captured["folder"] = folder

        def resolve(self):
            return result

    return FakeSingleFolderQuerySet


def test_calendar_folder_uses_mailbox_account_when_given(settings, monkeypatch) -> None:
    backend = EWSExchangeBackend(settings)
    other_calendar = SimpleNamespace(folder_class="IPF.Appointment")
    other_account = _CalendarOnlyAccount()
    monkeypatch.setattr(backend, "_account_for", lambda mailbox: other_account)
    captured: dict = {}
    monkeypatch.setattr(
        exchange_client_calendar,
        "SingleFolderQuerySet",
        _fake_single_folder_query_set(other_calendar, captured),
    )

    folder = backend._calendar_folder(None, "colleague@example.com")

    assert folder is other_calendar
    assert captured["account"] is other_account
    assert captured["folder"].id == "calendar"
    assert captured["folder"].mailbox.email_address == "colleague@example.com"


def test_calendar_folder_maps_access_denied_on_shared_calendar(settings, monkeypatch) -> None:
    backend = EWSExchangeBackend(settings)
    monkeypatch.setattr(backend, "_account_for", lambda mailbox: _CalendarOnlyAccount())
    monkeypatch.setattr(
        exchange_client_calendar,
        "SingleFolderQuerySet",
        _fake_single_folder_query_set(ErrorAccessDenied("denied"), {}),
    )

    with pytest.raises(PermissionDeniedError):
        backend._calendar_folder(None, "colleague@example.com")


def test_list_events_queries_the_mailbox_calendar_not_the_service_account(
    settings, monkeypatch
) -> None:
    backend = EWSExchangeBackend(settings)

    def own_view(**kwargs):
        raise AssertionError("must not query the service account's own calendar")

    backend._account = SimpleNamespace(
        calendar=SimpleNamespace(view=own_view), default_timezone=UTC
    )

    class _EventQuery(list):
        def only(self, *fields):
            return self

    other_calendar = SimpleNamespace(view=lambda start, end: _EventQuery())
    monkeypatch.setattr(
        backend, "_account_for", lambda mailbox: SimpleNamespace(default_timezone=UTC)
    )
    monkeypatch.setattr(backend, "_shared_calendar_folder", lambda mailbox: other_calendar)

    request = ListEventsRequest(
        start=datetime(2026, 4, 13, 9, 0, tzinfo=UTC),
        end=datetime(2026, 4, 13, 10, 0, tzinfo=UTC),
        mailbox="colleague@example.com",
    )
    result = backend.list_events(request)

    assert result == []


def test_get_event_fetches_from_the_mailbox_account(settings, monkeypatch) -> None:
    backend = EWSExchangeBackend(settings)
    other_calendar = SimpleNamespace(folder_class="IPF.Appointment")
    other_account = _CalendarOnlyAccount()
    monkeypatch.setattr(backend, "_account_for", lambda mailbox: other_account)
    monkeypatch.setattr(backend, "_shared_calendar_folder", lambda mailbox: other_calendar)

    captured: dict = {}

    def fake_fetch_item(item_id, folder=None, expected_type=None, account=None):
        captured["account"] = account
        captured["folder"] = folder
        return SimpleNamespace(
            id=item_id,
            subject="Sync",
            start=datetime(2026, 4, 13, 9, 0, tzinfo=UTC),
            end=datetime(2026, 4, 13, 9, 30, tzinfo=UTC),
        )

    monkeypatch.setattr(backend, "_fetch_item", fake_fetch_item)

    result = backend.get_event(GetEventRequest(id="event-1", mailbox="colleague@example.com"))

    assert captured["account"] is other_account
    assert captured["folder"] is other_calendar
    assert result.id == "event-1"


def test_mailbox_cannot_be_combined_with_calendar_id_for_list_events() -> None:
    with pytest.raises(ValidationError):
        ListEventsRequest(
            start=datetime(2026, 4, 13, 9, 0, tzinfo=UTC),
            end=datetime(2026, 4, 13, 10, 0, tzinfo=UTC),
            mailbox="colleague@example.com",
            calendar_id="some-folder-id",
        )


def test_mailbox_cannot_be_combined_with_calendar_id_for_get_event() -> None:
    with pytest.raises(ValidationError):
        GetEventRequest(id="event-1", mailbox="colleague@example.com", calendar_id="some-folder-id")
