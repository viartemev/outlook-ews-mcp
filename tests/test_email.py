from __future__ import annotations

from pathlib import Path

import pytest

from outlook_mcp.errors import APIError
from outlook_mcp.tools.email import (
    add_attachment,
    get_attachment,
    get_email,
    list_emails,
    send_email,
)


def test_list_emails(client) -> None:
    result = list_emails(client, {"folder": "inbox", "limit": 5, "unread_only": True})
    assert result[0]["id"] == "email-1"
    assert result[0]["is_read"] is False


def test_get_email(client) -> None:
    result = get_email(client, {"id": "email-1"})
    assert result["id"] == "email-1"
    assert result["body_text"] == "Body"


def test_send_email_checks_attachment_exists(client, tmp_path: Path) -> None:
    client.settings.attachment_root = tmp_path
    existing = tmp_path / "ok.txt"
    existing.write_text("hello", encoding="utf-8")

    result = send_email(
        client,
        {
            "to": ["user@example.com"],
            "subject": "Test",
            "body": "Hello",
            "attachments": [str(existing)],
        },
    )
    assert result["status"] == "sent"


def test_send_email_rejects_missing_attachment(client, tmp_path: Path) -> None:
    client.settings.attachment_root = tmp_path

    with pytest.raises(APIError) as excinfo:
        send_email(
            client,
            {
                "to": ["user@example.com"],
                "subject": "Test",
                "body": "Hello",
                "attachments": [str(tmp_path / "no-such-file.txt")],
            },
        )

    assert excinfo.value.code == "validation_error"


def test_add_attachment_rejects_paths_when_root_not_configured(client, tmp_path: Path) -> None:
    existing = tmp_path / "ok.txt"
    existing.write_text("hello", encoding="utf-8")

    with pytest.raises(APIError) as excinfo:
        add_attachment(client, {"email_id": "email-1", "path": str(existing)})

    assert excinfo.value.code == "validation_error"
    assert "EXCHANGE_ATTACHMENT_ROOT" in excinfo.value.details[0]["reason"]


def test_send_email_rejects_attachments_when_root_not_configured(client, tmp_path: Path) -> None:
    existing = tmp_path / "ok.txt"
    existing.write_text("hello", encoding="utf-8")

    with pytest.raises(APIError) as excinfo:
        send_email(
            client,
            {
                "to": ["user@example.com"],
                "subject": "Test",
                "body": "Hello",
                "attachments": [str(existing)],
            },
        )

    assert excinfo.value.code == "validation_error"
    assert "EXCHANGE_ATTACHMENT_ROOT" in excinfo.value.details[0]["reason"]


def test_send_email_allows_no_attachments_when_root_not_configured(client) -> None:
    result = send_email(
        client,
        {"to": ["user@example.com"], "subject": "Test", "body": "Hello"},
    )
    assert result["status"] == "sent"


def test_send_email_rejects_non_regular_file_attachment(client, tmp_path: Path) -> None:
    client.settings.attachment_root = tmp_path
    directory_as_attachment = tmp_path / "dir"
    directory_as_attachment.mkdir()

    with pytest.raises(APIError) as excinfo:
        send_email(
            client,
            {
                "to": ["user@example.com"],
                "subject": "Test",
                "body": "Hello",
                "attachments": [str(directory_as_attachment)],
            },
        )

    assert excinfo.value.code == "validation_error"
    assert "not a regular file" in excinfo.value.details[0]["reason"]


def test_send_email_rejects_attachment_over_limit(client, tmp_path: Path) -> None:
    client.settings.attachment_root = tmp_path
    client.settings.attachment_max_size_mb = 1
    oversized = tmp_path / "big.bin"
    oversized.write_bytes(b"x" * (1024 * 1024 + 1))

    with pytest.raises(APIError) as excinfo:
        send_email(
            client,
            {
                "to": ["user@example.com"],
                "subject": "Test",
                "body": "Hello",
                "attachments": [str(oversized)],
            },
        )

    assert excinfo.value.code == "validation_error"
    assert "EXCHANGE_ATTACHMENT_MAX_SIZE_MB=1" in excinfo.value.details[0]["reason"]


def test_send_email_rejects_too_many_attachments(client, tmp_path: Path) -> None:
    client.settings.attachment_root = tmp_path
    client.settings.attachment_max_count = 2
    paths = []
    for i in range(3):
        path = tmp_path / f"f{i}.txt"
        path.write_text("hi", encoding="utf-8")
        paths.append(str(path))

    with pytest.raises(APIError) as excinfo:
        send_email(
            client,
            {
                "to": ["user@example.com"],
                "subject": "Test",
                "body": "Hello",
                "attachments": paths,
            },
        )

    assert excinfo.value.code == "validation_error"
    assert "EXCHANGE_ATTACHMENT_MAX_COUNT=2" in excinfo.value.details[0]["reason"]


def test_send_email_rejects_total_attachment_size_over_limit(client, tmp_path: Path) -> None:
    client.settings.attachment_root = tmp_path
    client.settings.attachment_max_total_size_mb = 1
    paths = []
    for i in range(2):
        path = tmp_path / f"f{i}.bin"
        path.write_bytes(b"x" * (600 * 1024))
        paths.append(str(path))

    with pytest.raises(APIError) as excinfo:
        send_email(
            client,
            {
                "to": ["user@example.com"],
                "subject": "Test",
                "body": "Hello",
                "attachments": paths,
            },
        )

    assert excinfo.value.code == "validation_error"
    assert "EXCHANGE_ATTACHMENT_MAX_TOTAL_SIZE_MB=1" in excinfo.value.details[0]["reason"]


def test_send_email_rejects_attachment_outside_root(client, tmp_path: Path) -> None:
    root = tmp_path / "allowed"
    root.mkdir()
    client.settings.attachment_root = root
    outside = tmp_path / "outside.txt"
    outside.write_text("hello", encoding="utf-8")

    with pytest.raises(APIError) as excinfo:
        send_email(
            client,
            {
                "to": ["user@example.com"],
                "subject": "Test",
                "body": "Hello",
                "attachments": [str(outside)],
            },
        )

    assert excinfo.value.code == "validation_error"
    assert "EXCHANGE_ATTACHMENT_ROOT" in excinfo.value.details[0]["reason"]


def test_send_email_allows_attachment_inside_root(client, tmp_path: Path) -> None:
    root = tmp_path / "allowed"
    root.mkdir()
    client.settings.attachment_root = root
    inside = root / "ok.txt"
    inside.write_text("hello", encoding="utf-8")

    result = send_email(
        client,
        {
            "to": ["user@example.com"],
            "subject": "Test",
            "body": "Hello",
            "attachments": [str(inside)],
        },
    )
    assert result["status"] == "sent"


def test_get_attachment_rejects_save_path_outside_root(client, tmp_path: Path) -> None:
    root = tmp_path / "allowed"
    root.mkdir()
    client.settings.attachment_root = root
    outside = tmp_path / "elsewhere"

    with pytest.raises(APIError) as excinfo:
        get_attachment(
            client,
            {
                "email_id": "email-1",
                "attachment_id": "attachment-1",
                "save_path": str(outside),
            },
        )

    assert excinfo.value.code == "validation_error"
    assert "EXCHANGE_ATTACHMENT_ROOT" in excinfo.value.details[0]["reason"]


def test_get_attachment_allows_save_path_inside_root(client, tmp_path: Path) -> None:
    root = tmp_path / "allowed"
    root.mkdir()
    client.settings.attachment_root = root
    inside = root / "downloads"

    result = get_attachment(
        client,
        {
            "email_id": "email-1",
            "attachment_id": "attachment-1",
            "save_path": str(inside),
        },
    )
    assert result["filename"] == "test.txt"


def test_get_attachment_rejects_save_path_when_root_not_configured(client, tmp_path: Path) -> None:
    with pytest.raises(APIError) as excinfo:
        get_attachment(
            client,
            {
                "email_id": "email-1",
                "attachment_id": "attachment-1",
                "save_path": str(tmp_path / "downloads"),
            },
        )

    assert excinfo.value.code == "validation_error"
    assert "EXCHANGE_ATTACHMENT_ROOT" in excinfo.value.details[0]["reason"]


def test_get_attachment_allows_no_save_path_when_root_not_configured(client) -> None:
    result = get_attachment(client, {"email_id": "email-1", "attachment_id": "attachment-1"})
    assert result["filename"] == "test.txt"
