from __future__ import annotations

import anyio
import pytest

from outlook_mcp.mcp_tools import ToolSpec, bind_mcp_tool
from outlook_mcp.models import EmailSummary, ListEmailsRequest, SearchContactsRequest
from outlook_mcp.server import build_mcp_server, build_registry


@pytest.mark.parametrize(
    "model_name", ["UpdateDraftRequest", "UpdateContactRequest", "UpdateEventRequest"]
)
@pytest.mark.parametrize("extra", [{}, {"body": None}])
def test_fastmcp_preserves_omission(model_name, extra) -> None:
    from mcp.server.fastmcp.tools.base import Tool
    from outlook_mcp import models

    model = getattr(models, model_name)
    if model_name == "UpdateContactRequest" and extra:
        extra = {"company": None}
    captured = []

    def call(name, arguments):
        captured.append(arguments)
        return {}, False

    spec = ToolSpec("patch", "Patch", lambda c, a: None, request_model=model)
    tool = Tool.from_function(bind_mcp_tool(call, spec))
    arguments = {"id": "item-1", **extra}
    anyio.run(lambda: tool.run(arguments))
    assert captured == [arguments]


def test_update_draft_through_fastmcp_preserves_other_fields(client, settings) -> None:
    server = build_mcp_server(settings=settings, client=client)
    tool = server._tool_manager._tools["update_draft"]
    result = anyio.run(lambda: tool.run({"id": "draft-1", "to": ["a@example.com"]}))
    assert not result.isError
    assert result.structuredContent["updated_fields"] == ["to"]


def test_create_reply_draft_is_registered_without_send(client, settings) -> None:
    server = build_mcp_server(settings=settings, client=client)
    tool = server._tool_manager._tools["create_reply_draft"]
    result = anyio.run(lambda: tool.run({"id": "source", "body": "Reply"}))
    assert not result.isError
    assert result.structuredContent["status"] == "draft"
    assert tool.annotations.destructiveHint is False


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


def test_bind_mcp_tool_publishes_field_constraints(client, settings) -> None:
    """Regression guard: the dynamic tool signature used to carry only
    field.annotation, so ge/le/min_length constraints on the request model
    never reached the published MCP tool schema."""
    from outlook_mcp.models import SendEmailRequest

    registry = build_registry(settings=settings, client=client)
    spec = ToolSpec(
        "list_emails",
        "List emails in a folder",
        lambda client, arguments: None,
        request_model=ListEmailsRequest,
    )
    tool_fn = bind_mcp_tool(registry.call, spec)

    from mcp.server.fastmcp.tools.base import Tool

    tool = Tool.from_function(tool_fn, name="list_emails", description="List emails in a folder")
    limit_schema = tool.parameters["properties"]["limit"]

    assert limit_schema["minimum"] == 1
    assert limit_schema["maximum"] == 100

    send_spec = ToolSpec(
        "send_email",
        "Send an email",
        lambda client, arguments: None,
        request_model=SendEmailRequest,
    )
    send_tool_fn = bind_mcp_tool(registry.call, send_spec)
    send_tool = Tool.from_function(send_tool_fn, name="send_email", description="Send an email")

    assert send_tool.parameters["properties"]["to"]["minItems"] == 1


def test_bind_mcp_tool_routes_flat_arguments(client, settings) -> None:
    registry = build_registry(settings=settings, client=client)
    spec = ToolSpec(
        "search_contacts",
        "Search contacts",
        lambda client, arguments: None,
        request_model=SearchContactsRequest,
    )
    tool_fn = bind_mcp_tool(registry.call, spec)
    # Tool functions are coroutines so FastMCP awaits them instead of running
    # blocking Exchange work on the event loop thread.
    result = anyio.run(lambda: tool_fn(query="ivan"))
    assert result.isError is False
    assert result.structuredContent["result"][0]["id"] == "contact-1"


def test_build_mcp_server_registers_typed_tools(client, settings) -> None:
    server = build_mcp_server(settings=settings, client=client)
    list_emails_tool = server._tool_manager._tools["list_emails"]

    assert "folder" in list_emails_tool.parameters["properties"]
    assert "kwargs" not in list_emails_tool.parameters["properties"]


def test_state_replacing_tools_are_marked_destructive(client, settings) -> None:
    server = build_mcp_server(settings=settings, client=client)

    for name in [
        "send_email",
        "reply_email",
        "forward_email",
        "move_email",
        "mark_email",
        "send_draft",
        "create_event",
        "update_event",
        "respond_to_invite",
        "update_contact",
    ]:
        assert server._tool_manager._tools[name].annotations.destructiveHint is True

    for name in ["list_emails", "copy_email", "create_folder", "create_contact"]:
        assert server._tool_manager._tools[name].annotations.destructiveHint is False


def test_registry_rejects_legacy_kwargs_wrapper(client, settings) -> None:
    registry = build_registry(settings=settings, client=client)
    payload, is_error = registry.call("search_contacts", {"kwargs": {"query": "ivan"}})
    assert is_error is True
    assert payload["error"] == "validation_error"
