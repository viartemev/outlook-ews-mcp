from __future__ import annotations

from outlook_mcp.errors import APIError, NotFoundError
from outlook_mcp.exchange_client import EWSExchangeBackend
from outlook_mcp.exchange_client.base import BaseEWSBackend
from outlook_mcp.models import (
    ActionResult,
    BulkCategorizeEmailsRequest,
    BulkDeleteEmailsRequest,
    BulkMarkEmailsRequest,
    BulkMoveEmailsRequest,
)


def test_bulk_helper_captures_per_item_failure_without_aborting() -> None:
    backend = BaseEWSBackend.__new__(BaseEWSBackend)

    def action(item_id: str) -> ActionResult:
        if item_id == "bad":
            raise NotFoundError(item_id)
        return ActionResult(id=item_id, status="updated")

    results = backend._bulk(["good-1", "bad", "good-2"], action)

    assert results == [
        ActionResult(id="good-1", status="updated"),
        ActionResult(id="bad", status="error", warning="item bad was not found"),
        ActionResult(id="good-2", status="updated"),
    ]


def test_bulk_helper_lets_non_api_errors_propagate() -> None:
    backend = BaseEWSBackend.__new__(BaseEWSBackend)

    def action(item_id: str) -> ActionResult:
        raise ValueError("boom")

    try:
        backend._bulk(["x"], action)
    except ValueError:
        pass
    else:
        raise AssertionError("expected ValueError to propagate")


def test_bulk_move_emails_delegates_to_move_email_per_id(monkeypatch) -> None:
    backend = EWSExchangeBackend.__new__(EWSExchangeBackend)
    calls: list[str] = []

    def fake_move_email(request):
        calls.append(request.id)
        if request.id == "email-2":
            raise APIError("permission_denied", "no access")
        return ActionResult(id=request.id, status="moved", new_folder=request.folder)

    monkeypatch.setattr(backend, "move_email", fake_move_email)

    result = backend.bulk_move_emails(
        BulkMoveEmailsRequest(ids=["email-1", "email-2"], folder="archive")
    )

    assert calls == ["email-1", "email-2"]
    assert result[0] == ActionResult(id="email-1", status="moved", new_folder="archive")
    assert result[1].status == "error"
    assert result[1].id == "email-2"


def test_bulk_delete_emails_delegates_to_delete_email_per_id(monkeypatch) -> None:
    backend = EWSExchangeBackend.__new__(EWSExchangeBackend)
    calls: list[tuple[str, bool]] = []

    def fake_delete_email(request):
        calls.append((request.id, request.hard_delete))
        return ActionResult(id=request.id, status="deleted")

    monkeypatch.setattr(backend, "delete_email", fake_delete_email)

    result = backend.bulk_delete_emails(BulkDeleteEmailsRequest(ids=["a", "b"], hard_delete=True))

    assert calls == [("a", True), ("b", True)]
    assert [r.status for r in result] == ["deleted", "deleted"]


def test_bulk_mark_emails_delegates_to_mark_email_per_id(monkeypatch) -> None:
    backend = EWSExchangeBackend.__new__(EWSExchangeBackend)
    calls: list[tuple[str, bool | None]] = []

    def fake_mark_email(request):
        calls.append((request.id, request.read))
        return ActionResult(id=request.id, status="updated", updated_fields=["read"])

    monkeypatch.setattr(backend, "mark_email", fake_mark_email)

    result = backend.bulk_mark_emails(BulkMarkEmailsRequest(ids=["a", "b"], read=True))

    assert calls == [("a", True), ("b", True)]
    assert len(result) == 2


def test_bulk_categorize_emails_delegates_to_categorize_email_per_id(monkeypatch) -> None:
    backend = EWSExchangeBackend.__new__(EWSExchangeBackend)
    calls: list[tuple[str, list[str], str]] = []

    def fake_categorize_email(request):
        calls.append((request.id, request.categories, request.mode))
        return ActionResult(id=request.id, status="updated", categories=request.categories)

    monkeypatch.setattr(backend, "categorize_email", fake_categorize_email)

    result = backend.bulk_categorize_emails(
        BulkCategorizeEmailsRequest(ids=["a", "b"], categories=["Important"], mode="add")
    )

    assert calls == [("a", ["Important"], "add"), ("b", ["Important"], "add")]
    assert len(result) == 2
