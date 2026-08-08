from __future__ import annotations

from outlook_mcp.smoke import _env_flag, _mask_email, _sanitize_mailbox, _sanitize_ping


def test_mask_email() -> None:
    assert _mask_email("user@example.com") == "us**@example.com"
    assert _mask_email("ab@example.com") == "a*@example.com"
    assert _mask_email(None) is None


def test_sanitize_mailbox_masks_email_and_drops_identifying_fields() -> None:
    payload = {
        "email_address": "user@example.com",
        "display_name": "Test User",
        "timezone": "Europe/Moscow",
        "exchange_version": "2019",
    }

    result = _sanitize_mailbox(payload)

    assert result == {"email_address": "us**@example.com"}


def test_sanitize_ping_redacts_server() -> None:
    payload = {
        "status": "ok",
        "server": "mail.internal.company.com",
        "version": "Exchange2019_CU15",
        "latency_ms": 12,
    }

    result = _sanitize_ping(payload)

    assert result["server"] == "<redacted>"
    assert result["status"] == "ok"
    assert result["latency_ms"] == 12
    assert "version" not in result


def test_env_flag(monkeypatch) -> None:
    monkeypatch.setenv("OUTLOOK_MCP_SMOKE_INCLUDE_DATA", "true")
    assert _env_flag("OUTLOOK_MCP_SMOKE_INCLUDE_DATA") is True

    monkeypatch.setenv("OUTLOOK_MCP_SMOKE_INCLUDE_DATA", "0")
    assert _env_flag("OUTLOOK_MCP_SMOKE_INCLUDE_DATA") is False


def test_main_prints_a_sanitized_payload(monkeypatch, capsys, settings) -> None:
    import json

    import outlook_mcp.smoke as smoke
    from conftest import FakeExchangeBackend

    monkeypatch.setattr(smoke, "get_settings", lambda: settings)
    monkeypatch.setattr(smoke, "build_default_backend", lambda s: FakeExchangeBackend())
    monkeypatch.delenv("OUTLOOK_MCP_SMOKE_INCLUDE_DATA", raising=False)

    smoke.main()

    payload = json.loads(capsys.readouterr().out)
    assert payload["data_included"] is False
    assert "recent_inbox" not in payload
    # The mailbox address must come out masked, never verbatim.
    assert payload["mailbox_info"]["email_address"] != "user@example.com"
    assert payload["mailbox_info"]["email_address"].endswith("@example.com")


def test_main_includes_real_data_only_on_explicit_opt_in(monkeypatch, capsys, settings) -> None:
    import json

    import outlook_mcp.smoke as smoke
    from conftest import FakeExchangeBackend

    monkeypatch.setattr(smoke, "get_settings", lambda: settings)
    monkeypatch.setattr(smoke, "build_default_backend", lambda s: FakeExchangeBackend())
    monkeypatch.setenv("OUTLOOK_MCP_SMOKE_INCLUDE_DATA", "true")

    smoke.main()

    payload = json.loads(capsys.readouterr().out)
    assert payload["data_included"] is True
    assert payload["recent_inbox"]
    assert payload["upcoming_events"]
