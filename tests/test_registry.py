from __future__ import annotations

from outlook_mcp.server import build_registry


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


def test_registry_resource_mailbox_email(client, settings) -> None:
    registry = build_registry(settings=settings, client=client)
    payload, is_error = registry.call("resource_mailbox_email", {"id": "email-1"})
    assert is_error is False
    assert payload["data"]["id"] == "email-1"


def test_registry_calendar_today_resource(client, settings) -> None:
    registry = build_registry(settings=settings, client=client)
    payload, is_error = registry.call("resource_calendar_today", {})
    assert is_error is False
    assert payload["uri"] == "calendar://today"
