from __future__ import annotations

from . import models
from .mcp_tools import ToolSpec
from .tools.calendar import (
    create_event,
    delete_events,
    find_free_slots,
    get_event,
    get_my_availability,
    list_calendars,
    list_events,
    list_room_lists,
    list_rooms,
    respond_to_invites,
    update_event,
)
from .tools.contacts import (
    create_contact,
    delete_contact,
    get_contact,
    search_contacts,
    update_contact,
)
from .tools.email import (
    categorize_emails,
    create_draft,
    create_reply_draft,
    create_folder,
    delete_emails,
    delete_folder,
    delete_attachment,
    forward_email,
    get_attachment,
    get_email,
    get_email_mime,
    get_thread,
    add_attachment,
    list_categories,
    list_emails,
    list_folders,
    mark_emails,
    move_emails,
    copy_emails,
    rename_folder,
    reply_email,
    search_emails,
    send_draft,
    send_email,
    update_draft,
)
from .tools.system import (
    create_inbox_rule,
    delete_inbox_rule,
    get_mailbox_info,
    get_out_of_office,
    list_delegates,
    list_inbox_rules,
    ping_exchange,
    update_inbox_rule,
    set_out_of_office,
)

#: The single source of truth for every Outlook MCP tool: name, description,
#: handler, request/response schema, and MCP annotations all live together so
#: adding a tool means adding one entry here instead of touching N call sites.
TOOL_SPECS: list[ToolSpec] = [
    ToolSpec(
        "ping_exchange",
        "Check connectivity to Exchange",
        ping_exchange,
        response_model=models.PingResult,
        read_only=True,
    ),
    ToolSpec(
        "get_mailbox_info",
        "Get the mailbox's address, display name, timezone and Exchange version",
        get_mailbox_info,
        response_model=models.MailboxInfo,
        read_only=True,
    ),
    ToolSpec(
        "list_delegates",
        "List mailbox delegates and their folder permission levels. Read-only: "
        "managing delegates is not supported by exchangelib.",
        list_delegates,
        response_model=list[models.DelegateInfo],
        read_only=True,
    ),
    ToolSpec(
        "list_inbox_rules",
        "List server-side inbox rules (the curated subset of conditions and "
        "actions this API models)",
        list_inbox_rules,
        response_model=list[models.InboxRule],
        read_only=True,
    ),
    ToolSpec(
        "create_inbox_rule",
        "Create a server-side inbox rule, e.g. 'from this sender -> move to "
        "folder'. WARNING: managing rules over EWS removes the client-side rule blob desktop Outlook keeps, which can wipe rules created in Outlook itself.",
        create_inbox_rule,
        request_model=models.CreateInboxRuleRequest,
        response_model=models.InboxRule,
        destructive=True,
    ),
    ToolSpec(
        "update_inbox_rule",
        "Enable/disable an inbox rule or change its priority. Other fields are "
        "deliberately not updatable here. WARNING: managing rules over EWS removes the client-side rule blob desktop Outlook keeps, which can wipe rules created in Outlook itself.",
        update_inbox_rule,
        request_model=models.UpdateInboxRuleRequest,
        response_model=models.InboxRule,
        destructive=True,
    ),
    ToolSpec(
        "delete_inbox_rule",
        "Delete a server-side inbox rule by id. WARNING: managing rules over EWS removes the client-side rule blob desktop Outlook keeps, which can wipe rules created in Outlook itself.",
        delete_inbox_rule,
        request_model=models.DeleteInboxRuleRequest,
        response_model=models.ActionResult,
        destructive=True,
    ),
    ToolSpec(
        "get_out_of_office",
        "Get the mailbox out-of-office (automatic reply) settings",
        get_out_of_office,
        response_model=models.OutOfOfficeSettings,
        read_only=True,
    ),
    ToolSpec(
        "set_out_of_office",
        "Set the mailbox out-of-office (automatic reply): disabled, enabled, or "
        "scheduled with a start/end window",
        set_out_of_office,
        request_model=models.OutOfOfficeSettings,
        response_model=models.OutOfOfficeSettings,
        destructive=True,
    ),
    ToolSpec(
        "list_emails",
        "List emails in a folder (default inbox), newest first, with paging (limit/offset) "
        "and filters: sender, subject substring, date range, unread, has attachments. Returns "
        "summaries; use get_email or get_thread for full bodies",
        list_emails,
        request_model=models.ListEmailsRequest,
        response_model=list[models.EmailSummary],
        read_only=True,
    ),
    ToolSpec(
        "get_email",
        "Get one email by id with recipients, text and HTML body (truncated past "
        "EXCHANGE_EMAIL_BODY_MAX_CHARS) and attachment metadata. Set include_headers for the "
        "raw internet headers",
        get_email,
        request_model=models.GetEmailRequest,
        response_model=models.EmailFull,
        read_only=True,
    ),
    ToolSpec(
        "get_email_mime",
        "Export a message's raw RFC 822 MIME content, base64-encoded",
        get_email_mime,
        request_model=models.GetEmailMimeRequest,
        response_model=models.EmailMimeResult,
        read_only=True,
    ),
    ToolSpec(
        "get_thread",
        "Get every message of a conversation in chronological order, bodies "
        "included. Pass the id of any message in the thread, or a conversation_id "
        "from a listing.",
        get_thread,
        request_model=models.GetThreadRequest,
        response_model=models.Thread,
        read_only=True,
    ),
    ToolSpec(
        "search_emails",
        "Search emails. Pass `query` for a substring match over subject/body/sender, "
        "or `aqs` for server-side Advanced Query Syntax "
        "(e.g. 'from:ivan AND hasattachments:true').",
        search_emails,
        request_model=models.SearchEmailsRequest,
        response_model=list[models.EmailSummary],
        read_only=True,
    ),
    ToolSpec(
        "send_email",
        "Send a new email. Set from_mailbox to send from a shared mailbox you have "
        "Send As rights on; its Sent Items keeps the copy",
        send_email,
        request_model=models.SendEmailRequest,
        response_model=models.SendResult,
        destructive=True,
    ),
    ToolSpec(
        "reply_email",
        "Reply to an email and send it immediately: to the sender only, or to everyone with "
        "reply_all. Supports extra To/CC recipients and attachments. Use create_reply_draft "
        "to review before sending",
        reply_email,
        request_model=models.ReplyEmailRequest,
        response_model=models.SendResult,
        destructive=True,
    ),
    ToolSpec(
        "forward_email",
        "Forward an email to new recipients and send it immediately, with an optional comment "
        "above the original and extra attachments",
        forward_email,
        request_model=models.ForwardEmailRequest,
        response_model=models.SendResult,
        destructive=True,
    ),
    ToolSpec(
        "move_emails",
        "Move one or more emails (ids) to another folder in one call. Per-item results: one "
        "bad id does not fail the rest. Moved items get new ids, returned as new_id; the old "
        "ids stop working",
        move_emails,
        request_model=models.BulkMoveEmailsRequest,
        response_model=models.BulkResult,
        destructive=True,
    ),
    ToolSpec(
        "copy_emails",
        "Copy one or more emails (ids) to another folder in one call, with per-item results; "
        "each copy's id is returned as new_id",
        copy_emails,
        request_model=models.BulkMoveEmailsRequest,
        response_model=models.BulkResult,
    ),
    ToolSpec(
        "delete_emails",
        "Delete one or more emails (ids) in one call, with per-item results. Moves them to "
        "Deleted Items (recoverable) unless hard_delete is set, which removes them "
        "permanently",
        delete_emails,
        request_model=models.BulkDeleteEmailsRequest,
        response_model=models.BulkResult,
        destructive=True,
    ),
    ToolSpec(
        "mark_emails",
        "Set read/unread, importance, or the follow-up flag (optionally with start and due "
        "dates) on one or more emails (ids), with per-item results. Omitted fields are left "
        "unchanged",
        mark_emails,
        request_model=models.BulkMarkEmailsRequest,
        response_model=models.BulkResult,
        destructive=True,
    ),
    ToolSpec(
        "categorize_emails",
        "Set, add, or remove Outlook categories (the coloured labels) on one or more emails "
        "(ids), with per-item results. mode=set with an empty list clears all categories",
        categorize_emails,
        request_model=models.BulkCategorizeEmailsRequest,
        response_model=models.BulkResult,
        destructive=True,
    ),
    ToolSpec(
        "list_categories",
        "List the categories in use, with counts. Collected from the most recent "
        "messages of the given folders, not from the mailbox master category list.",
        list_categories,
        request_model=models.ListCategoriesRequest,
        response_model=list[models.CategoryUsage],
        read_only=True,
    ),
    ToolSpec(
        "list_folders",
        "List mailbox folders as a tree with unread/total counts, starting under parent "
        "(default: the mailbox root) down to depth levels",
        list_folders,
        request_model=models.ListFoldersRequest,
        response_model=list[models.FolderInfo],
        read_only=True,
    ),
    ToolSpec(
        "create_folder",
        "Create a mailbox folder under parent (default inbox)",
        create_folder,
        request_model=models.CreateFolderRequest,
        response_model=models.ActionResult,
    ),
    ToolSpec(
        "rename_folder",
        "Rename a mailbox folder. Refuses distinguished folders (Inbox, Sent Items, "
        "Calendar, etc.)",
        rename_folder,
        request_model=models.RenameFolderRequest,
        response_model=models.ActionResult,
        destructive=True,
    ),
    ToolSpec(
        "delete_folder",
        "Delete a mailbox folder and everything in it. Soft-deletes to Deleted Items "
        "unless hard_delete is set. Refuses distinguished folders (Inbox, Sent Items, "
        "Calendar, etc.)",
        delete_folder,
        request_model=models.DeleteFolderRequest,
        response_model=models.ActionResult,
        destructive=True,
    ),
    ToolSpec(
        "create_draft",
        "Save a new email to Drafts without sending it. Returns the draft id for "
        "update_draft, add_attachment and send_draft",
        create_draft,
        request_model=models.DraftEmailRequest,
        response_model=models.ActionResult,
    ),
    ToolSpec(
        "create_reply_draft",
        "Create a reply-all draft in the original thread without sending. Supports HTML, "
        "{{mention:key}} tokens, additional To/CC recipients and attachments.",
        create_reply_draft,
        request_model=models.CreateReplyDraftRequest,
        response_model=models.ActionResult,
    ),
    ToolSpec(
        "update_draft",
        "Update an existing email draft. Omitted fields are left unchanged; "
        "'attachments', if given, replaces the draft's entire attachment set",
        update_draft,
        request_model=models.UpdateDraftRequest,
        response_model=models.ActionResult,
        destructive=True,
    ),
    ToolSpec(
        "send_draft",
        "Submit a draft and verify its copy in Sent Items. sent_confirmed means the copy "
        "was found, not recipient delivery. submitted/submitted_unconfirmed must not be "
        "retried blindly. Returns the actual sent-copy mailbox and correlation ID.",
        send_draft,
        request_model=models.SendDraftRequest,
        response_model=models.SendResult,
        destructive=True,
    ),
    ToolSpec(
        "add_attachment",
        "Attach a local file to an existing message, typically a draft. The file "
        "must live under EXCHANGE_ATTACHMENT_ROOT.",
        add_attachment,
        request_model=models.AddAttachmentRequest,
        response_model=models.ActionResult,
        destructive=True,
    ),
    ToolSpec(
        "delete_attachment",
        "Remove one attachment from a message by attachment id",
        delete_attachment,
        request_model=models.DeleteAttachmentRequest,
        response_model=models.ActionResult,
        destructive=True,
    ),
    ToolSpec(
        "get_attachment",
        "Download one attachment of an email to local disk (under EXCHANGE_ATTACHMENT_ROOT "
        "when save_path is given, else a temp directory) and return the saved path. Never "
        "overwrites an existing file",
        get_attachment,
        request_model=models.GetAttachmentRequest,
        response_model=models.AttachmentResult,
        read_only=True,
    ),
    ToolSpec(
        "list_events",
        "List up to limit calendar events in a time range. Set mailbox to view a colleague's "
        "default calendar instead of your own (requires read access to that calendar "
        "already granted on the server; not combinable with calendar_id). If only "
        "free/busy is shared, events come from the availability service with an empty "
        "id and only time, status, subject and location",
        list_events,
        request_model=models.ListEventsRequest,
        response_model=list[models.CalendarEvent],
        read_only=True,
    ),
    ToolSpec(
        "get_event",
        "Get a calendar event by ID. Set mailbox to read from a colleague's "
        "default calendar instead of your own",
        get_event,
        request_model=models.GetEventRequest,
        response_model=models.CalendarEvent,
        read_only=True,
    ),
    ToolSpec(
        "create_event",
        "Create a calendar event, optionally all-day or recurring. If attendees are given "
        "this is a meeting and invitations are sent to them immediately",
        create_event,
        request_model=models.CreateEventRequest,
        response_model=models.CreateEventResult,
        destructive=True,
    ),
    ToolSpec(
        "update_event",
        "Update a calendar event: subject, time, location, body, reminder, add/remove "
        "attendees. Omitted fields are left unchanged. send_updates (default all) controls "
        "whether attendees are notified",
        update_event,
        request_model=models.UpdateEventRequest,
        response_model=models.ActionResult,
        destructive=True,
    ),
    ToolSpec(
        "find_free_slots",
        "Find up to limit meeting time slots",
        find_free_slots,
        request_model=models.FindFreeSlotsRequest,
        response_model=list[models.FreeSlot],
        read_only=True,
    ),
    ToolSpec(
        "delete_events",
        "Delete one or more calendar events (ids), with per-item results. For meetings you "
        "organize, attendees get a cancellation (with cancel_message) unless notify_attendees "
        "is false",
        delete_events,
        request_model=models.BulkDeleteEventsRequest,
        response_model=models.BulkResult,
        destructive=True,
    ),
    ToolSpec(
        "respond_to_invites",
        "Accept, tentatively accept, or decline one or more meeting invitations (ids) with an "
        "optional message; the organizer is notified. Per-item results",
        respond_to_invites,
        request_model=models.BulkRespondToInvitesRequest,
        response_model=models.BulkResult,
        destructive=True,
    ),
    ToolSpec(
        "get_my_availability",
        "Get free and busy slots. Set mailbox to check a colleague's calendar instead of your own",
        get_my_availability,
        request_model=models.ListEventsRequest,
        response_model=models.AvailabilityResult,
        read_only=True,
    ),
    ToolSpec(
        "list_calendars",
        "List the calendar folders in your own mailbox; their ids can be passed as calendar_id",
        list_calendars,
        response_model=list[models.CalendarInfo],
        read_only=True,
    ),
    ToolSpec(
        "list_room_lists",
        "List Room Finder room lists (groups of meeting rooms)",
        list_room_lists,
        response_model=list[models.RoomListInfo],
        read_only=True,
    ),
    ToolSpec(
        "list_rooms",
        "List the meeting rooms in a Room Finder room list",
        list_rooms,
        request_model=models.ListRoomsRequest,
        response_model=list[models.RoomInfo],
        read_only=True,
    ),
    ToolSpec(
        "search_contacts",
        "Search contacts by name or email: personal Contacts folder, the global address list "
        "(GAL), or both (default)",
        search_contacts,
        request_model=models.SearchContactsRequest,
        response_model=list[models.ContactSummary],
        read_only=True,
    ),
    ToolSpec(
        "get_contact",
        "Get a contact by id, including its email addresses, phones, addresses and notes",
        get_contact,
        request_model=models.GetContactRequest,
        response_model=models.ContactFull,
        read_only=True,
    ),
    ToolSpec(
        "create_contact",
        "Create a contact in your personal Contacts folder",
        create_contact,
        request_model=models.CreateContactRequest,
        response_model=models.ActionResult,
    ),
    ToolSpec(
        "update_contact",
        "Update a personal contact. Omitted fields are left unchanged; a field set to null is "
        "cleared",
        update_contact,
        request_model=models.UpdateContactRequest,
        response_model=models.ActionResult,
        destructive=True,
    ),
    ToolSpec(
        "delete_contact",
        "Delete a personal contact. Moves it to Deleted Items (recoverable) unless "
        "hard_delete is set",
        delete_contact,
        request_model=models.DeleteContactRequest,
        response_model=models.ActionResult,
        destructive=True,
    ),
]
