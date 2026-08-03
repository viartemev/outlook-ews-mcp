from __future__ import annotations

from outlook_mcp.mcp_tools import ToolSpec, bind_mcp_tool, normalize_tool_arguments
from outlook_mcp.models import EmailSummary, ListEmailsRequest, SearchContactsRequest
from outlook_mcp.server import build_mcp_server, build_registry


def test_normalize_tool_arguments_unwraps_legacy_kwargs() -> None:
    assert normalize_tool_arguments({"kwargs": {"folder": "inbox", "limit": 5}}) == {
        "folder": "inbox",
        "limit": 5,
    }


def test_normalize_tool_arguments_keeps_flat_payload() -> None:
    assert normalize_tool_arguments({"folder": "inbox"}) == {"folder": "inbox"}


def test_bind_mcp_tool_exposes_request_schema(client, settings) -> None:
    registry = build_registry(settings=settings, client=client)
    spec = ToolSpec(
        "list_emails",
        "List emails in a folder",
        lambda client, arguments: None,
        request_model=ListEmailsRequest,
        response_model=list[EmailSummary],
    )
    tool_fn = bind_mcp_tool(registry.call, spec)

    from mcp.server.fastmcp.tools.base import Tool

    tool = Tool.from_function(tool_fn, name="list_emails", description="List emails in a folder")
    schema = tool.parameters

    assert "properties" in schema
    assert "folder" in schema["properties"]
    assert "limit" in schema["properties"]
    assert "kwargs" not in schema["properties"]


def test_bind_mcp_tool_routes_flat_arguments(client, settings) -> None:
    registry = build_registry(settings=settings, client=client)
    spec = ToolSpec(
        "search_contacts",
        "Search contacts",
        lambda client, arguments: None,
        request_model=SearchContactsRequest,
    )
    tool_fn = bind_mcp_tool(registry.call, spec)
    result = tool_fn(query="ivan")
    assert result.isError is False
    assert result.structuredContent["result"][0]["id"] == "contact-1"


def test_build_mcp_server_registers_typed_tools(client, settings) -> None:
    server = build_mcp_server(settings=settings, client=client)
    list_emails_tool = server._tool_manager._tools["list_emails"]

    assert "folder" in list_emails_tool.parameters["properties"]
    assert "kwargs" not in list_emails_tool.parameters["properties"]


def test_registry_accepts_legacy_kwargs_wrapper(client, settings) -> None:
    registry = build_registry(settings=settings, client=client)
    payload, is_error = registry.call("search_contacts", {"kwargs": {"query": "ivan"}})
    assert is_error is False
    assert payload[0]["id"] == "contact-1"
