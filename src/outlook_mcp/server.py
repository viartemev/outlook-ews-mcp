from __future__ import annotations

import json
import logging
import sys
import time
from collections.abc import Callable
from typing import Any, Literal

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

    def _call(name: str, args: dict) -> str:
        payload, is_error = registry.call(name, args)
        if is_error:
            raise RuntimeError(json.dumps(payload, ensure_ascii=False))
        return json.dumps(payload, ensure_ascii=False)

    # --- System ---

    @server.tool(name="ping_exchange", description="Check connectivity to Exchange")
    def _ping_exchange() -> str:
        return _call("ping_exchange", {})

    @server.tool(name="get_mailbox_info", description="Get mailbox metadata")
    def _get_mailbox_info() -> str:
        return _call("get_mailbox_info", {})

    # --- Email ---

    @server.tool(
        name="list_emails",
        description='List emails in a folder. folder: "inbox"|"sent"|"drafts"|"deleted" or custom path. since/before: YYYY-MM-DD.',
    )
    def _list_emails(
        folder: str = "inbox",
        limit: int = 20,
        offset: int = 0,
        from_address: str | None = None,
        subject: str | None = None,
        since: str | None = None,
        before: str | None = None,
        unread_only: bool = False,
        has_attachments: bool | None = None,
    ) -> str:
        return _call("list_emails", {
            "folder": folder, "limit": limit, "offset": offset,
            "from_address": from_address, "subject": subject,
            "since": since, "before": before,
            "unread_only": unread_only, "has_attachments": has_attachments,
        })

    @server.tool(name="get_email", description="Get a full email by its EWS item ID")
    def _get_email(id: str) -> str:
        return _call("get_email", {"id": id})

    @server.tool(
        name="search_emails",
        description="Search emails by full text, subject, or sender. folder: optional folder path to restrict search.",
    )
    def _search_emails(query: str, folder: str | None = None, limit: int = 20) -> str:
        return _call("search_emails", {"query": query, "folder": folder, "limit": limit})

    @server.tool(
        name="send_email",
        description='Send a new email. to/cc/bcc: lists of email addresses. body_type: "text"|"html". importance: "low"|"normal"|"high". attachments: local file paths.',
    )
    def _send_email(
        to: list[str],
        subject: str,
        body: str,
        body_type: Literal["text", "html"] = "text",
        cc: list[str] = [],
        bcc: list[str] = [],
        reply_to: str | None = None,
        attachments: list[str] = [],
        importance: Literal["low", "normal", "high"] = "normal",
    ) -> str:
        return _call("send_email", {
            "to": to, "subject": subject, "body": body, "body_type": body_type,
            "cc": cc, "bcc": bcc, "reply_to": reply_to,
            "attachments": attachments, "importance": importance,
        })

    @server.tool(name="reply_email", description="Reply to an email. reply_all: reply to all recipients. attachments: local file paths.")
    def _reply_email(
        id: str,
        body: str,
        reply_all: bool = False,
        attachments: list[str] = [],
    ) -> str:
        return _call("reply_email", {"id": id, "body": body, "reply_all": reply_all, "attachments": attachments})

    @server.tool(name="forward_email", description="Forward an email to new recipients. attachments: additional local file paths.")
    def _forward_email(
        id: str,
        to: list[str],
        comment: str | None = None,
        attachments: list[str] = [],
    ) -> str:
        return _call("forward_email", {"id": id, "to": to, "comment": comment, "attachments": attachments})

    @server.tool(
        name="move_email",
        description=(
            'Move an email to another folder. '
            'For built-in folders use: "inbox", "sent", "drafts", "deleted", "junk", "archive". '
            'For subfolders use the full path as returned by list_folders, e.g. "inbox/_Claude/Support". '
            'If the target folder is not a built-in, call list_folders first to get the exact path.'
        ),
    )
    def _move_email(id: str, folder: str) -> str:
        return _call("move_email", {"id": id, "folder": folder})

    @server.tool(
        name="copy_email",
        description=(
            'Copy an email to another folder. '
            'For built-in folders use: "inbox", "sent", "drafts", "deleted", "junk", "archive". '
            'For subfolders use the full path as returned by list_folders, e.g. "inbox/_Claude/Support". '
            'If the target folder is not a built-in, call list_folders first to get the exact path.'
        ),
    )
    def _copy_email(id: str, folder: str) -> str:
        return _call("copy_email", {"id": id, "folder": folder})

    @server.tool(
        name="delete_email",
        description="Delete an email. hard_delete=true permanently deletes; default moves to Deleted Items.",
    )
    def _delete_email(id: str, hard_delete: bool = False) -> str:
        return _call("delete_email", {"id": id, "hard_delete": hard_delete})

    @server.tool(
        name="mark_email",
        description='Update email flags. read: true/false. flag: "flagged"|"complete"|"none". importance: "low"|"normal"|"high". At least one field required.',
    )
    def _mark_email(
        id: str,
        read: bool | None = None,
        flag: Literal["flagged", "complete", "none"] | None = None,
        importance: Literal["low", "normal", "high"] | None = None,
    ) -> str:
        return _call("mark_email", {"id": id, "read": read, "flag": flag, "importance": importance})

    @server.tool(name="list_folders", description="List mailbox folder tree. depth: nesting levels to return (default 2). The path field of each folder can be used directly as the folder parameter in move_email and copy_email.")
    def _list_folders(parent: str | None = None, depth: int = 2) -> str:
        return _call("list_folders", {"parent": parent, "depth": depth})

    @server.tool(name="create_folder", description='Create a new mailbox folder. parent: parent folder path (default: "inbox").')
    def _create_folder(name: str, parent: str | None = "inbox") -> str:
        return _call("create_folder", {"name": name, "parent": parent})

    @server.tool(name="create_draft", description='Save an email as draft without sending. body_type: "text"|"html".')
    def _create_draft(
        to: list[str],
        subject: str,
        body: str,
        body_type: Literal["text", "html"] = "text",
        cc: list[str] = [],
        bcc: list[str] = [],
        attachments: list[str] = [],
    ) -> str:
        return _call("create_draft", {
            "to": to, "subject": subject, "body": body, "body_type": body_type,
            "cc": cc, "bcc": bcc, "attachments": attachments,
        })

    @server.tool(name="send_draft", description="Send a previously saved draft email by its ID.")
    def _send_draft(id: str) -> str:
        return _call("send_draft", {"id": id})

    @server.tool(
        name="get_attachment",
        description="Save an email attachment to disk. save_path: optional destination file path (default: temp directory).",
    )
    def _get_attachment(email_id: str, attachment_id: str, save_path: str | None = None) -> str:
        return _call("get_attachment", {"email_id": email_id, "attachment_id": attachment_id, "save_path": save_path})

    # --- Calendar ---

    @server.tool(
        name="list_events",
        description="List calendar events in a time range. start/end: ISO 8601 datetime. calendar_id: optional, use list_calendars to get IDs.",
    )
    def _list_events(
        start: str,
        end: str,
        calendar_id: str | None = None,
        include_recurring: bool = True,
    ) -> str:
        return _call("list_events", {
            "start": start, "end": end,
            "calendar_id": calendar_id, "include_recurring": include_recurring,
        })

    @server.tool(name="get_event", description="Get a calendar event by its ID.")
    def _get_event(id: str) -> str:
        return _call("get_event", {"id": id})

    @server.tool(
        name="create_event",
        description=(
            'Create a calendar event. start/end: ISO 8601 datetime. attendees: list of email addresses. '
            'recurrence: {"type":"daily"|"weekly"|"monthly"|"yearly","interval":1,"end_date":"YYYY-MM-DD"}. '
            'importance: "low"|"normal"|"high".'
        ),
    )
    def _create_event(
        subject: str,
        start: str,
        end: str,
        calendar_id: str | None = None,
        location: str | None = None,
        body: str | None = None,
        attendees: list[str] = [],
        is_all_day: bool = False,
        reminder_minutes: int | None = 15,
        recurrence: dict | None = None,
        categories: list[str] = [],
        importance: Literal["low", "normal", "high"] = "normal",
        online_meeting: bool = False,
    ) -> str:
        return _call("create_event", {
            "subject": subject, "start": start, "end": end,
            "calendar_id": calendar_id, "location": location, "body": body,
            "attendees": attendees, "is_all_day": is_all_day,
            "reminder_minutes": reminder_minutes, "recurrence": recurrence,
            "categories": categories, "importance": importance, "online_meeting": online_meeting,
        })

    @server.tool(
        name="update_event",
        description='Update a calendar event. Only provide fields to change. send_updates: "none"|"all"|"modified".',
    )
    def _update_event(
        id: str,
        subject: str | None = None,
        start: str | None = None,
        end: str | None = None,
        location: str | None = None,
        body: str | None = None,
        add_attendees: list[str] = [],
        remove_attendees: list[str] = [],
        reminder_minutes: int | None = None,
        send_updates: Literal["none", "all", "modified"] = "all",
    ) -> str:
        return _call("update_event", {
            "id": id, "subject": subject, "start": start, "end": end,
            "location": location, "body": body,
            "add_attendees": add_attendees, "remove_attendees": remove_attendees,
            "reminder_minutes": reminder_minutes, "send_updates": send_updates,
        })

    @server.tool(
        name="delete_event",
        description="Delete a calendar event. notify_attendees: send cancellation notice (default true).",
    )
    def _delete_event(
        id: str,
        notify_attendees: bool = True,
        cancel_message: str | None = None,
    ) -> str:
        return _call("delete_event", {"id": id, "notify_attendees": notify_attendees, "cancel_message": cancel_message})

    @server.tool(
        name="respond_to_invite",
        description='Respond to a calendar invite. response: "accept"|"tentative"|"decline". message: optional reply text.',
    )
    def _respond_to_invite(
        id: str,
        response: Literal["accept", "tentative", "decline"],
        message: str | None = None,
    ) -> str:
        return _call("respond_to_invite", {"id": id, "response": response, "message": message})

    @server.tool(
        name="find_free_slots",
        description='Find available meeting time slots. attendees: email addresses. duration: minutes. start/end: ISO 8601. work_hours: {"start":"09:00","end":"18:00"}.',
    )
    def _find_free_slots(
        attendees: list[str],
        duration: int,
        start: str,
        end: str,
        work_hours: dict | None = None,
    ) -> str:
        return _call("find_free_slots", {
            "attendees": attendees, "duration": duration,
            "start": start, "end": end, "work_hours": work_hours,
        })

    @server.tool(name="get_my_availability", description="Get my free and busy time slots. start/end: ISO 8601 datetime.")
    def _get_my_availability(
        start: str,
        end: str,
        calendar_id: str | None = None,
        include_recurring: bool = True,
    ) -> str:
        return _call("get_my_availability", {
            "start": start, "end": end,
            "calendar_id": calendar_id, "include_recurring": include_recurring,
        })

    @server.tool(name="list_calendars", description="List all available calendars. Use returned id values as calendar_id in other tools.")
    def _list_calendars() -> str:
        return _call("list_calendars", {})

    # --- Contacts ---

    @server.tool(
        name="search_contacts",
        description='Search contacts by name, email, company, or job title. source: "personal"|"gal"|"all".',
    )
    def _search_contacts(
        query: str,
        source: Literal["personal", "gal", "all"] = "all",
        limit: int = 10,
    ) -> str:
        return _call("search_contacts", {"query": query, "source": source, "limit": limit})

    @server.tool(name="get_contact", description="Get full contact details by ID.")
    def _get_contact(id: str) -> str:
        return _call("get_contact", {"id": id})

    @server.tool(name="create_contact", description="Create a new personal contact.")
    def _create_contact(
        display_name: str,
        first_name: str | None = None,
        last_name: str | None = None,
        email: str | None = None,
        phone: str | None = None,
        company: str | None = None,
        job_title: str | None = None,
        notes: str | None = None,
    ) -> str:
        return _call("create_contact", {
            "display_name": display_name, "first_name": first_name, "last_name": last_name,
            "email": email, "phone": phone, "company": company, "job_title": job_title, "notes": notes,
        })

    @server.tool(name="update_contact", description="Update an existing personal contact. display_name is required; all other fields are optional.")
    def _update_contact(
        id: str,
        display_name: str,
        first_name: str | None = None,
        last_name: str | None = None,
        email: str | None = None,
        phone: str | None = None,
        company: str | None = None,
        job_title: str | None = None,
        notes: str | None = None,
    ) -> str:
        return _call("update_contact", {
            "id": id, "display_name": display_name, "first_name": first_name, "last_name": last_name,
            "email": email, "phone": phone, "company": company, "job_title": job_title, "notes": notes,
        })

    @server.tool(name="delete_contact", description="Delete a personal contact by its ID.")
    def _delete_contact(id: str) -> str:
        return _call("delete_contact", {"id": id})

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
