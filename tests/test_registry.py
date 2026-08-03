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
