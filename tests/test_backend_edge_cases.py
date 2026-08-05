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
    def __init__(
        self, attachment_id: str, name: str, content: bytes, size: int | None = None
    ) -> None:
        self.attachment_id = SimpleNamespace(id=attachment_id)
        self.name = name
        self.content = content
        self.content_type = "text/plain"
        self.size = size


def test_get_attachment_sanitizes_relative_traversal(settings, tmp_path: Path, monkeypatch) -> None:
    backend = EWSExchangeBackend(settings)
    attachment = FakeAttachment("att-1", "../../evil.txt", b"payload")
    monkeypatch.setattr(
        backend, "_fetch_item", lambda *a, **k: SimpleNamespace(attachments=[attachment])
    )

    request = GetAttachmentRequest(
        email_id="email-1", attachment_id="att-1", save_path=str(tmp_path)
    )
    result = backend.get_attachment(request)

    saved_path = Path(result.saved_path)
    assert saved_path.parent == tmp_path
    assert saved_path.name == "evil.txt"


def test_get_attachment_sanitizes_absolute_path(settings, tmp_path: Path, monkeypatch) -> None:
    backend = EWSExchangeBackend(settings)
    outside_target = tmp_path.parent / "outside-target.txt"
    attachment = FakeAttachment("att-1", str(outside_target), b"payload")
    monkeypatch.setattr(
        backend, "_fetch_item", lambda *a, **k: SimpleNamespace(attachments=[attachment])
    )

    request = GetAttachmentRequest(
        email_id="email-1", attachment_id="att-1", save_path=str(tmp_path)
    )
    result = backend.get_attachment(request)

    saved_path = Path(result.saved_path)
    assert saved_path.parent == tmp_path
    assert not outside_target.exists()


def test_get_attachment_does_not_follow_dangling_symlink(
    settings, tmp_path: Path, monkeypatch
) -> None:
    backend = EWSExchangeBackend(settings)
    outside_target = tmp_path.parent / "outside-via-symlink.txt"
    (tmp_path / "note.txt").symlink_to(outside_target)
    attachment = FakeAttachment("att-1", "note.txt", b"payload")
    monkeypatch.setattr(
        backend, "_fetch_item", lambda *a, **k: SimpleNamespace(attachments=[attachment])
    )

    result = backend.get_attachment(
        GetAttachmentRequest(email_id="email-1", attachment_id="att-1", save_path=str(tmp_path))
    )

    assert Path(result.saved_path).name == "note-1.txt"
    assert Path(result.saved_path).read_bytes() == b"payload"
    assert not outside_target.exists()


def test_get_attachment_rejects_declared_size_over_limit(
    settings, tmp_path: Path, monkeypatch
) -> None:
    settings.attachment_max_size_mb = 1
    backend = EWSExchangeBackend(settings)
    attachment = FakeAttachment("att-1", "big.bin", b"x", size=2 * 1024 * 1024)
    monkeypatch.setattr(
        backend, "_fetch_item", lambda *a, **k: SimpleNamespace(attachments=[attachment])
    )

    request = GetAttachmentRequest(
        email_id="email-1", attachment_id="att-1", save_path=str(tmp_path)
    )
    with pytest.raises(APIError) as excinfo:
        backend.get_attachment(request)

    assert excinfo.value.code == "validation_error"
    assert list(tmp_path.iterdir()) == []


def test_get_attachment_rejects_content_over_limit_when_size_unknown(
    settings, tmp_path: Path, monkeypatch
) -> None:
    settings.attachment_max_size_mb = 1
    backend = EWSExchangeBackend(settings)
    attachment = FakeAttachment("att-1", "big.bin", b"x" * (2 * 1024 * 1024))
    monkeypatch.setattr(
        backend, "_fetch_item", lambda *a, **k: SimpleNamespace(attachments=[attachment])
    )

    request = GetAttachmentRequest(
        email_id="email-1", attachment_id="att-1", save_path=str(tmp_path)
    )
    with pytest.raises(APIError) as excinfo:
        backend.get_attachment(request)

    assert excinfo.value.code == "validation_error"
    assert list(tmp_path.iterdir()) == []


class FakeStreamingAttachment:
    """Mimics exchangelib's FileAttachment.fp: a context manager streaming content
    in chunks, the way the real GetAttachment service does."""

    def __init__(
        self, attachment_id: str, name: str, content: bytes, size: int | None = None
    ) -> None:
        self.attachment_id = SimpleNamespace(id=attachment_id)
        self.name = name
        self.content_type = "application/octet-stream"
        self.size = size if size is not None else len(content)
        self._content = content

    class _Reader:
        def __init__(self, content: bytes) -> None:
            self._chunks = [content[i : i + 4] for i in range(0, len(content), 4)]

        def read(self, _size: int) -> bytes:
            return self._chunks.pop(0) if self._chunks else b""

    @property
    def fp(self):
        return self

    def __enter__(self):
        return self._Reader(self._content)

    def __exit__(self, *args):
        return False


def test_get_attachment_streams_via_fp_instead_of_content(
    settings, tmp_path: Path, monkeypatch
) -> None:
    backend = EWSExchangeBackend(settings)
    attachment = FakeStreamingAttachment("att-1", "note.txt", b"hello world, streamed")
    monkeypatch.setattr(
        backend, "_fetch_item", lambda *a, **k: SimpleNamespace(attachments=[attachment])
    )

    request = GetAttachmentRequest(
        email_id="email-1", attachment_id="att-1", save_path=str(tmp_path)
    )
    result = backend.get_attachment(request)

    assert Path(result.saved_path).read_bytes() == b"hello world, streamed"
    assert result.size == len(b"hello world, streamed")


def test_get_attachment_aborts_stream_over_limit_and_discards_partial_file(
    settings, tmp_path: Path, monkeypatch
) -> None:
    """The size check must run as chunks arrive, not only after the whole
    attachment has been buffered -- and a rejected download must not leave a
    partial file behind."""
    settings.attachment_max_size_mb = 1
    backend = EWSExchangeBackend(settings)
    attachment = FakeStreamingAttachment("att-1", "big.bin", b"x" * (2 * 1024 * 1024), size=None)
    monkeypatch.setattr(
        backend, "_fetch_item", lambda *a, **k: SimpleNamespace(attachments=[attachment])
    )

    request = GetAttachmentRequest(
        email_id="email-1", attachment_id="att-1", save_path=str(tmp_path)
    )
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


def test_get_my_availability_ignores_free_status_events(settings, monkeypatch) -> None:
    backend = EWSExchangeBackend(settings)
    backend._account = SimpleNamespace(default_timezone=EWSTimeZone("Europe/Moscow"))

    free_event = CalendarEvent(
        id="event-1",
        subject="Holiday",
        start=datetime(2026, 4, 8, 9, 0, tzinfo=UTC),
        end=datetime(2026, 4, 8, 12, 0, tzinfo=UTC),
        organizer=EmailAddress(email="organizer@example.com"),
        free_busy_status="free",
    )
    monkeypatch.setattr(backend, "list_events", lambda request: [free_event])

    request = ListEventsRequest(
        start=datetime(2026, 4, 8, 9, 0, tzinfo=UTC),
        end=datetime(2026, 4, 8, 12, 0, tzinfo=UTC),
    )
    result = backend.get_my_availability(request)

    assert result.busy_slots == []
    assert len(result.free_slots) == 1
    assert result.free_slots[0].start.timestamp() == request.start.timestamp()
    assert result.free_slots[0].end.timestamp() == request.end.timestamp()


def test_to_calendar_event_normalizes_free_busy_status(settings) -> None:
    backend = EWSExchangeBackend(settings)
    item = SimpleNamespace(
        id="event-1",
        subject="Sync",
        start=datetime(2026, 4, 8, 9, 0, tzinfo=UTC),
        end=datetime(2026, 4, 8, 10, 0, tzinfo=UTC),
        organizer=None,
        required_attendees=[],
        optional_attendees=[],
        legacy_free_busy_status="Free",
    )

    event = backend._to_calendar_event(item)

    assert event.free_busy_status == "free"
