from __future__ import annotations

from datetime import UTC, datetime

import pytest

from outlook_mcp.errors import APIError, NotFoundError
from outlook_mcp.exchange_client import EWSExchangeBackend
from outlook_mcp.exchange_client.base import BaseEWSBackend
from outlook_mcp.models import (
    ActionResult,
    BulkCategorizeEmailsRequest,
    BulkDeleteEventsRequest,
    BulkItemFailure,
    BulkItemResult,
    BulkMarkEmailsRequest,
    BulkRespondToInvitesRequest,
    BulkResult,
)


def test_bulk_helper_captures_per_item_failure_without_aborting() -> None:
    backend = BaseEWSBackend.__new__(BaseEWSBackend)

    def action(item_id: str) -> None:
        if item_id == "bad":
            raise NotFoundError(item_id)

    result = backend._bulk(["good-1", "bad", "good-2"], action)

    assert result == BulkResult(
        succeeded=[BulkItemResult(id="good-1"), BulkItemResult(id="good-2")],
        failed=[BulkItemFailure(id="bad", error="not_found", message="item bad was not found")],
    )


def test_bulk_helper_lets_non_api_errors_propagate() -> None:
    backend = BaseEWSBackend.__new__(BaseEWSBackend)

    def action(item_id: str) -> None:
        raise ValueError("boom")

    try:
        backend._bulk(["x"], action)
    except ValueError:
        pass
    else:
        raise AssertionError("expected ValueError to propagate")


def test_mark_emails_delegates_to_mark_email_per_id(monkeypatch) -> None:
    backend = EWSExchangeBackend.__new__(EWSExchangeBackend)
    calls: list[tuple[str, bool | None]] = []

    def fake_mark_email(request):
        calls.append((request.id, request.read))
        if request.id == "email-2":
            raise APIError("permission_denied", "no access")
        return ActionResult(id=request.id, status="updated")

    monkeypatch.setattr(backend, "mark_email", fake_mark_email)

    result = backend.mark_emails(BulkMarkEmailsRequest(ids=["email-1", "email-2"], read=True))

    assert calls == [("email-1", True), ("email-2", True)]
    assert result.succeeded == [BulkItemResult(id="email-1")]
    assert result.failed == [
        BulkItemFailure(id="email-2", error="permission_denied", message="no access")
    ]


def test_categorize_emails_delegates_to_categorize_email_per_id(monkeypatch) -> None:
    backend = EWSExchangeBackend.__new__(EWSExchangeBackend)
    calls: list[tuple[str, list[str], str]] = []

    def fake_categorize_email(request):
        calls.append((request.id, request.categories, request.mode))
        return ActionResult(id=request.id, status="updated", categories=request.categories)

    monkeypatch.setattr(backend, "categorize_email", fake_categorize_email)

    result = backend.categorize_emails(
        BulkCategorizeEmailsRequest(ids=["a", "b"], categories=["Important"], mode="add")
    )

    assert calls == [("a", ["Important"], "add"), ("b", ["Important"], "add")]
    assert [item.id for item in result.succeeded] == ["a", "b"]


def test_delete_events_delegates_to_delete_event_per_id(monkeypatch) -> None:
    backend = EWSExchangeBackend.__new__(EWSExchangeBackend)
    calls: list[tuple[str, bool]] = []

    def fake_delete_event(request):
        calls.append((request.id, request.notify_attendees))
        if request.id == "event-2":
            raise NotFoundError(request.id)
        return ActionResult(id=request.id, status="deleted")

    monkeypatch.setattr(backend, "delete_event", fake_delete_event)

    result = backend.delete_events(
        BulkDeleteEventsRequest(ids=["event-1", "event-2"], notify_attendees=False)
    )

    assert calls == [("event-1", False), ("event-2", False)]
    assert result.succeeded == [BulkItemResult(id="event-1")]
    assert result.failed == [
        BulkItemFailure(id="event-2", error="not_found", message="item event-2 was not found")
    ]


def test_respond_to_invites_delegates_to_respond_to_invite_per_id(monkeypatch) -> None:
    backend = EWSExchangeBackend.__new__(EWSExchangeBackend)
    calls: list[tuple[str, str]] = []

    def fake_respond_to_invite(request):
        calls.append((request.id, request.response))
        return ActionResult(id=request.id, status=request.response)

    monkeypatch.setattr(backend, "respond_to_invite", fake_respond_to_invite)

    result = backend.respond_to_invites(
        BulkRespondToInvitesRequest(ids=["event-1", "event-2"], response="accept")
    )

    assert calls == [("event-1", "accept"), ("event-2", "accept")]
    assert [item.id for item in result.succeeded] == ["event-1", "event-2"]


def test_bulk_mark_emails_rejects_due_date_before_start_date() -> None:
    with pytest.raises(Exception):
        BulkMarkEmailsRequest(
            ids=["a"],
            flag_start_date=datetime(2026, 4, 2, tzinfo=UTC),
            flag_due_date=datetime(2026, 4, 1, tzinfo=UTC),
        )


def test_bulk_mark_emails_rejects_dates_combined_with_flag_none() -> None:
    with pytest.raises(Exception):
        BulkMarkEmailsRequest(
            ids=["a"], flag="none", flag_start_date=datetime(2026, 4, 1, tzinfo=UTC)
        )


def test_bulk_categorize_emails_rejects_empty_categories_for_add() -> None:
    with pytest.raises(Exception):
        BulkCategorizeEmailsRequest(ids=["a"], mode="add", categories=[])


def test_bulk_categorize_emails_rejects_blank_category_names() -> None:
    with pytest.raises(Exception):
        BulkCategorizeEmailsRequest(ids=["a"], categories=["  "])
