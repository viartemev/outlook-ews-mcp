from __future__ import annotations

import json

import pytest
from mcp.shared.memory import create_connected_server_and_client_session

from outlook_mcp.server import build_mcp_server

pytestmark = pytest.mark.anyio


async def test_call_tool_returns_structured_content(client, settings) -> None:
    server = build_mcp_server(settings=settings, client=client)
    async with create_connected_server_and_client_session(server) as session:
        result = await session.call_tool("search_contacts", {"query": "ivan"})

    assert result.isError is False
    assert result.structuredContent is not None
    assert result.structuredContent["result"][0]["id"] == "contact-1"

    assert result.content
    payload = json.loads(result.content[0].text)
    assert payload["id"] == "contact-1"


async def test_call_tool_reports_is_error_for_cross_field_validation(client, settings) -> None:
    """``end <= start`` only fails ExchangeModel's own validator, not FastMCP's
    per-field arguments schema, so this exercises our internal error mapping."""
    server = build_mcp_server(settings=settings, client=client)
    async with create_connected_server_and_client_session(server) as session:
        result = await session.call_tool(
            "list_events",
            {"start": "2026-04-08T10:00:00Z", "end": "2026-04-08T09:00:00Z"},
        )

    assert result.isError is True
    assert result.content
    assert "validation_error" in result.content[0].text
