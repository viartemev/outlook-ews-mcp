from __future__ import annotations

from datetime import UTC, datetime
from pathlib import Path
from types import SimpleNamespace

import pytest
from exchangelib.ewsdatetime import EWSTimeZone

from outlook_mcp.errors import APIError
from outlook_mcp.exchange_client import EWSExchangeBackend
from outlook_mcp.models import CalendarEvent, EmailAddress, GetAttachmentRequest, ListEventsRequest


class FakeAttachment:
    def __init__(self, attachment_id: str, name: str, content: bytes, size: int | None = None) -> None:
        self.attachment_id = SimpleNamespace(id=attachment_id)
        self.name = name
        self.content = content
        self.content_type = "text/plain"
        self.size = size


def test_get_attachment_sanitizes_relative_traversal(settings, tmp_path: Path, monkeypatch) -> None:
    backend = EWSExchangeBackend(settings)
    attachment = FakeAttachment("att-1", "../../evil.txt", b"payload")
    monkeypatch.setattr(backend, "_fetch_item", lambda *a, **k: SimpleNamespace(attachments=[attachment]))

    request = GetAttachmentRequest(email_id="email-1", attachment_id="att-1", save_path=str(tmp_path))
    result = backend.get_attachment(request)

    saved_path = Path(result.saved_path)
    assert saved_path.parent == tmp_path
    assert saved_path.name == "evil.txt"


def test_get_attachment_sanitizes_absolute_path(settings, tmp_path: Path, monkeypatch) -> None:
    backend = EWSExchangeBackend(settings)
    outside_target = tmp_path.parent / "outside-target.txt"
    attachment = FakeAttachment("att-1", str(outside_target), b"payload")
    monkeypatch.setattr(backend, "_fetch_item", lambda *a, **k: SimpleNamespace(attachments=[attachment]))

    request = GetAttachmentRequest(email_id="email-1", attachment_id="att-1", save_path=str(tmp_path))
    result = backend.get_attachment(request)

    saved_path = Path(result.saved_path)
    assert saved_path.parent == tmp_path
    assert not outside_target.exists()


def test_get_attachment_rejects_declared_size_over_limit(settings, tmp_path: Path, monkeypatch) -> None:
    settings.attachment_max_size_mb = 1
    backend = EWSExchangeBackend(settings)
    attachment = FakeAttachment("att-1", "big.bin", b"x", size=2 * 1024 * 1024)
    monkeypatch.setattr(backend, "_fetch_item", lambda *a, **k: SimpleNamespace(attachments=[attachment]))

    request = GetAttachmentRequest(email_id="email-1", attachment_id="att-1", save_path=str(tmp_path))
    with pytest.raises(APIError) as excinfo:
        backend.get_attachment(request)

    assert excinfo.value.code == "validation_error"
    assert list(tmp_path.iterdir()) == []


def test_get_attachment_rejects_content_over_limit_when_size_unknown(settings, tmp_path: Path, monkeypatch) -> None:
    settings.attachment_max_size_mb = 1
    backend = EWSExchangeBackend(settings)
    attachment = FakeAttachment("att-1", "big.bin", b"x" * (2 * 1024 * 1024))
    monkeypatch.setattr(backend, "_fetch_item", lambda *a, **k: SimpleNamespace(attachments=[attachment]))

    request = GetAttachmentRequest(email_id="email-1", attachment_id="att-1", save_path=str(tmp_path))
    with pytest.raises(APIError) as excinfo:
        backend.get_attachment(request)

    assert excinfo.value.code == "validation_error"
    assert list(tmp_path.iterdir()) == []


def test_get_my_availability_computes_free_slots(settings, monkeypatch) -> None:
    backend = EWSExchangeBackend(settings)
    backend._account = SimpleNamespace(default_timezone=EWSTimeZone("Europe/Moscow"))

    busy_event = CalendarEvent(
        id="event-1",
        subject="Sync",
        start=datetime(2026, 4, 8, 10, 0, tzinfo=UTC),
        end=datetime(2026, 4, 8, 10, 30, tzinfo=UTC),
        organizer=EmailAddress(email="organizer@example.com"),
    )
    monkeypatch.setattr(backend, "list_events", lambda request: [busy_event])

    request = ListEventsRequest(
        start=datetime(2026, 4, 8, 9, 0, tzinfo=UTC),
        end=datetime(2026, 4, 8, 12, 0, tzinfo=UTC),
    )
    result = backend.get_my_availability(request)

    assert len(result.busy_slots) == 1
    assert len(result.free_slots) == 2
    assert result.free_slots[0].start.timestamp() == request.start.timestamp()
    assert result.free_slots[0].end.timestamp() == busy_event.start.timestamp()
    assert result.free_slots[1].start.timestamp() == busy_event.end.timestamp()
    assert result.free_slots[1].end.timestamp() == request.end.timestamp()


def test_get_my_availability_fully_busy_has_no_free_slots(settings, monkeypatch) -> None:
    backend = EWSExchangeBackend(settings)
    backend._account = SimpleNamespace(default_timezone=EWSTimeZone("Europe/Moscow"))

    busy_event = CalendarEvent(
        id="event-1",
        subject="All day sync",
        start=datetime(2026, 4, 8, 9, 0, tzinfo=UTC),
        end=datetime(2026, 4, 8, 12, 0, tzinfo=UTC),
        organizer=EmailAddress(email="organizer@example.com"),
    )
    monkeypatch.setattr(backend, "list_events", lambda request: [busy_event])

    request = ListEventsRequest(
        start=datetime(2026, 4, 8, 9, 0, tzinfo=UTC),
        end=datetime(2026, 4, 8, 12, 0, tzinfo=UTC),
    )
    result = backend.get_my_availability(request)

    assert result.free_slots == []
