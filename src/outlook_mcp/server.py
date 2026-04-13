from __future__ import annotations

import json
import logging
import sys
import time
from collections.abc import Callable
from typing import Any

from .config import Settings, get_settings
from .errors import APIError, normalize_exception
from .exchange_client import ExchangeClient, build_default_backend
from .tools.calendar import (
    create_event,
    delete_event,
    find_free_slots,
    get_event,
    get_my_availability,
    list_calendars,
    list_events,
    respond_to_invite,
    update_event,
)
from .tools.contacts import create_contact, delete_contact, get_contact, search_contacts, update_contact
from .tools.email import (
    copy_email,
    create_draft,
    create_folder,
    delete_email,
    forward_email,
    get_attachment,
    get_email,
    list_emails,
    list_folders,
    mark_email,
    move_email,
    reply_email,
    search_emails,
    send_draft,
    send_email,
)
from .tools.system import get_mailbox_info, ping_exchange

logger = logging.getLogger(__name__)

ToolHandler = Callable[[ExchangeClient, dict[str, Any]], Any]


def configure_logging(settings: Settings) -> None:
    handlers: list[logging.Handler]
    if settings.log_file and str(settings.log_file).strip() not in {"", "."} and not settings.log_file.is_dir():
        handlers = [logging.FileHandler(settings.log_file)]
    else:
        handlers = [logging.StreamHandler(sys.stderr)]
    logging.basicConfig(
        level=getattr(logging, settings.log_level),
        format="%(asctime)s %(levelname)s %(name)s %(message)s",
        handlers=handlers,
    )


class ToolRegistry:
    def __init__(self, client: ExchangeClient) -> None:
        self.client = client
        self._handlers: dict[str, ToolHandler] = {
            "list_emails": list_emails,
            "get_email": get_email,
            "search_emails": search_emails,
            "send_email": send_email,
            "reply_email": reply_email,
            "forward_email": forward_email,
            "move_email": move_email,
            "copy_email": copy_email,
            "delete_email": delete_email,
            "mark_email": mark_email,
            "list_folders": list_folders,
            "create_folder": create_folder,
            "create_draft": create_draft,
            "send_draft": send_draft,
            "get_attachment": get_attachment,
            "list_events": list_events,
            "get_event": get_event,
            "create_event": create_event,
            "update_event": update_event,
            "delete_event": delete_event,
            "respond_to_invite": respond_to_invite,
            "find_free_slots": find_free_slots,
            "get_my_availability": get_my_availability,
            "search_contacts": search_contacts,
            "get_contact": get_contact,
            "create_contact": create_contact,
            "update_contact": update_contact,
            "delete_contact": delete_contact,
        }

    def call(self, name: str, arguments: dict[str, Any] | None = None) -> tuple[dict[str, Any], bool]:
        arguments = arguments or {}
        if name == "ping_exchange":
            return ping_exchange(self.client), False
        if name == "get_mailbox_info":
            return get_mailbox_info(self.client), False
        if name == "list_calendars":
            return list_calendars(self.client), False
        if name == "resource_mailbox_folders":
            return {"uri": "mailbox://folders", "data": list_folders(self.client, {})}, False
        if name == "resource_mailbox_inbox":
            return {
                "uri": "mailbox://inbox?limit=10",
                "data": list_emails(self.client, {"folder": "inbox", "limit": 10}),
            }, False
        if name == "resource_mailbox_email":
            email_id = arguments.get("id")
            return {"uri": f"mailbox://email/{email_id}", "data": get_email(self.client, arguments)}, False
        if name == "resource_mailbox_drafts":
            return {"uri": "mailbox://drafts", "data": list_emails(self.client, {"folder": "drafts"})}, False
        if name == "resource_calendar_today":
            now = time.time()
            return {"uri": "calendar://today", "generated_at": now}, False

        handler = self._handlers.get(name)
        if not handler:
            error = APIError("not_found", f"unknown tool: {name}")
            return error.to_dict(), True

        started = time.perf_counter()
        try:
            result = handler(self.client, arguments)
            duration_ms = round((time.perf_counter() - started) * 1000)
            logger.info("tool=%s status=ok duration_ms=%s", name, duration_ms)
            return result, False
        except Exception as exc:  # noqa: BLE001
            api_error = normalize_exception(exc)
            duration_ms = round((time.perf_counter() - started) * 1000)
            logger.warning(
                "tool=%s status=error duration_ms=%s error=%s",
                name,
                duration_ms,
                api_error.code,
            )
            return api_error.to_dict(), True


def build_registry(settings: Settings | None = None, client: ExchangeClient | None = None) -> ToolRegistry:
    settings = settings or get_settings()
    configure_logging(settings)
    client = client or ExchangeClient(settings=settings, backend=build_default_backend(settings))
    return ToolRegistry(client)


def build_mcp_server(settings: Settings | None = None, client: ExchangeClient | None = None) -> Any:
    try:
        from mcp.server.fastmcp import FastMCP
    except ImportError as exc:  # pragma: no cover
        raise RuntimeError("mcp package is required to run the server") from exc

    registry = build_registry(settings=settings, client=client)
    server = FastMCP("outlook-mcp")

    def register_tool(name: str, description: str) -> None:
        @server.tool(name=name, description=description)
        def _tool(**kwargs: Any) -> str:
            payload, is_error = registry.call(name, kwargs)
            if is_error:
                raise RuntimeError(json.dumps(payload, ensure_ascii=False))
            return json.dumps(payload, ensure_ascii=False)

    register_tool("ping_exchange", "Check connectivity to Exchange")
    register_tool("get_mailbox_info", "Get mailbox metadata")
    register_tool("list_emails", "List emails in a folder")
    register_tool("get_email", "Get a full email by ID")
    register_tool("search_emails", "Search emails")
    register_tool("send_email", "Send a new email")
    register_tool("reply_email", "Reply to an email")
    register_tool("forward_email", "Forward an email")
    register_tool("move_email", "Move email to another folder")
    register_tool("copy_email", "Copy email to another folder")
    register_tool("delete_email", "Delete an email")
    register_tool("mark_email", "Update email flags")
    register_tool("list_folders", "List mailbox folders")
    register_tool("create_folder", "Create a mailbox folder")
    register_tool("create_draft", "Create an email draft")
    register_tool("send_draft", "Send a draft email")
    register_tool("get_attachment", "Save an attachment to disk")
    register_tool("list_events", "List calendar events in a time range")
    register_tool("get_event", "Get a calendar event by ID")
    register_tool("create_event", "Create a calendar event")
    register_tool("update_event", "Update a calendar event")
    register_tool("delete_event", "Delete a calendar event")
    register_tool("respond_to_invite", "Respond to a calendar invite")
    register_tool("find_free_slots", "Find meeting time slots")
    register_tool("get_my_availability", "Get free and busy slots")
    register_tool("list_calendars", "List calendars")
    register_tool("search_contacts", "Search contacts")
    register_tool("get_contact", "Get a contact by ID")
    register_tool("create_contact", "Create a personal contact")
    register_tool("update_contact", "Update a personal contact")
    register_tool("delete_contact", "Delete a personal contact")
    return server


def main() -> None:
    settings = get_settings()
    server = build_mcp_server(settings=settings)
    transport = settings.mcp_transport

    if transport == "stdio":
        server.run()
        return

    if transport == "sse":  # pragma: no cover
        server.run(transport="sse", host=settings.mcp_sse_host, port=settings.mcp_sse_port)
        return

    raise RuntimeError(f"unsupported transport: {transport}")
