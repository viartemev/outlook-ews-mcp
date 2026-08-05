from __future__ import annotations

from dataclasses import replace

from outlook_mcp.server import build_registry
from outlook_mcp.models import SendResult
from outlook_mcp.tool_specs import TOOL_SPECS


def test_registry_knows_extended_tools(client, settings) -> None:
    registry = build_registry(settings=settings, client=client)

    payload, is_error = registry.call("search_contacts", {"query": "ivan"})
    assert is_error is False
    assert payload[0]["id"] == "contact-1"

    payload, is_error = registry.call(
        "list_calendars",
        {},
    )
    assert is_error is False
    assert payload[0]["is_default"] is True


def test_registry_handles_unknown_tool(client, settings) -> None:
    registry = build_registry(settings=settings, client=client)
    payload, is_error = registry.call("nope", {})
    assert is_error is True
    assert payload["error"] == "not_found"


def test_send_draft_registry_uses_send_result_schema() -> None:
    spec = next(spec for spec in TOOL_SPECS if spec.name == "send_draft")

    assert spec.response_model is SendResult


def test_registry_does_not_log_raw_unexpected_exception(client, settings, caplog) -> None:
    registry = build_registry(settings=settings, client=client)
    secret_detail = "mail.internal.example user@example.com"

    def fail(_client, _arguments):
        raise RuntimeError(secret_detail)

    registry._specs["ping_exchange"] = replace(registry._specs["ping_exchange"], handler=fail)
    payload, is_error = registry.call("ping_exchange")

    assert is_error is True
    assert payload["error"] == "internal_error"
    assert secret_detail not in caplog.text
    assert "exception_type=RuntimeError" in caplog.text


def test_registry_validates_success_payload_against_response_model(client, settings) -> None:
    registry = build_registry(settings=settings, client=client)

    def invalid_ping(_client, _arguments):
        return {"status": "broken"}

    registry._specs["ping_exchange"] = replace(
        registry._specs["ping_exchange"], handler=invalid_ping
    )
    payload, is_error = registry.call("ping_exchange")

    assert is_error is True
    assert payload["error"] == "internal_error"
