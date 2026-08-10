from __future__ import annotations

from outlook_mcp.errors import NotFoundError
from outlook_mcp.exchange_client import EWSExchangeBackend
from outlook_mcp.models import (
    ActionResult,
    BulkDeleteEventsRequest,
    BulkRespondToInvitesRequest,
)


def test_bulk_delete_events_delegates_to_delete_event_per_id(monkeypatch) -> None:
    backend = EWSExchangeBackend.__new__(EWSExchangeBackend)
    calls: list[tuple[str, bool]] = []

    def fake_delete_event(request):
        calls.append((request.id, request.notify_attendees))
        if request.id == "event-2":
            raise NotFoundError(request.id)
        return ActionResult(id=request.id, status="deleted")

    monkeypatch.setattr(backend, "delete_event", fake_delete_event)

    result = backend.bulk_delete_events(
        BulkDeleteEventsRequest(ids=["event-1", "event-2"], notify_attendees=False)
    )

    assert calls == [("event-1", False), ("event-2", False)]
    assert result[0] == ActionResult(id="event-1", status="deleted")
    assert result[1].status == "error"
    assert result[1].id == "event-2"


def test_bulk_respond_to_invites_delegates_to_respond_to_invite_per_id(monkeypatch) -> None:
    backend = EWSExchangeBackend.__new__(EWSExchangeBackend)
    calls: list[tuple[str, str]] = []

    def fake_respond_to_invite(request):
        calls.append((request.id, request.response))
        return ActionResult(id=request.id, status=request.response)

    monkeypatch.setattr(backend, "respond_to_invite", fake_respond_to_invite)

    result = backend.bulk_respond_to_invites(
        BulkRespondToInvitesRequest(ids=["event-1", "event-2"], response="accept")
    )

    assert calls == [("event-1", "accept"), ("event-2", "accept")]
    assert [r.status for r in result] == ["accept", "accept"]


def test_bulk_delete_events_propagates_non_api_errors(monkeypatch) -> None:
    backend = EWSExchangeBackend.__new__(EWSExchangeBackend)

    def fake_delete_event(request):
        raise ValueError("boom")

    monkeypatch.setattr(backend, "delete_event", fake_delete_event)

    try:
        backend.bulk_delete_events(BulkDeleteEventsRequest(ids=["event-1"]))
    except ValueError:
        pass
    else:
        raise AssertionError("expected ValueError to propagate")
