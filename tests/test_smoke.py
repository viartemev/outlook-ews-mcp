from __future__ import annotations

from outlook_mcp.smoke import _env_flag, _mask_email, _sanitize_mailbox


def test_mask_email() -> None:
    assert _mask_email("user@example.com") == "us**@example.com"
    assert _mask_email("ab@example.com") == "a*@example.com"
    assert _mask_email(None) is None


def test_sanitize_mailbox_masks_email() -> None:
    payload = {
        "email_address": "user@example.com",
        "display_name": "Test User",
        "timezone": "Europe/Moscow",
    }

    result = _sanitize_mailbox(payload)

    assert result["email_address"] == "us**@example.com"
    assert result["display_name"] == "Test User"


def test_env_flag(monkeypatch) -> None:
    monkeypatch.setenv("OUTLOOK_MCP_SMOKE_INCLUDE_DATA", "true")
    assert _env_flag("OUTLOOK_MCP_SMOKE_INCLUDE_DATA") is True

    monkeypatch.setenv("OUTLOOK_MCP_SMOKE_INCLUDE_DATA", "0")
    assert _env_flag("OUTLOOK_MCP_SMOKE_INCLUDE_DATA") is False
