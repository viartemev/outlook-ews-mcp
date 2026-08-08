from __future__ import annotations

import re
from collections.abc import Sequence
from datetime import date, datetime
from pathlib import Path
from typing import Annotated, Any, Literal

from pydantic import (
    BaseModel,
    BeforeValidator,
    ConfigDict,
    EmailStr,
    Field,
    field_validator,
    model_validator,
)

_HHMM_RE = re.compile(r"^([01]\d|2[0-3]):[0-5]\d$")

WeekdayName = Literal["monday", "tuesday", "wednesday", "thursday", "friday", "saturday", "sunday"]


def _server_address(value: Any) -> Any:
    """Normalise an address reported by Exchange for an item that already exists.

    Exchange does not always hand back SMTP. For internal senders, distribution
    lists, GAL entries and migrated mailboxes it uses routing type ``EX`` and
    returns a legacy X.500 distinguished name instead::

        /o=TANDER/ou=Exchange Administrative Group (FYDIBOHF23SPDLT)/cn=Recipients/cn=<hex>-<alias>

    Validating those as e-mail fails, and because a folder listing builds one
    model per message, a single such message used to blow up the whole call.
    Read paths must survive whatever the server reports, so the value is passed
    through unchanged.

    It is deliberately *not* rewritten into a plausible SMTP address: a
    fabricated address could later be used as a recipient and quietly deliver
    mail to the wrong person, or to nobody at all.
    """
    if isinstance(value, str):
        return value.strip()
    return value


#: Address as reported by the server. Lenient on purpose — see ``_server_address``.
#: Outgoing addresses supplied by the caller stay ``EmailStr``: garbage must be
#: rejected before it reaches Exchange, not after.
ServerAddress = Annotated[str, BeforeValidator(_server_address), Field(min_length=1)]


class ExchangeModel(BaseModel):
    model_config = ConfigDict(populate_by_name=True, extra="forbid")


class EmailAddress(ExchangeModel):
    email: ServerAddress
    name: str | None = None


class Attachment(ExchangeModel):
    id: str | None = None
    name: str
    size: int | None = None
    content_type: str | None = None
    # False for an embedded Exchange item (a forwarded email/calendar invite/contact
    # attached as an item rather than a file) -- get_attachment can only save file
    # attachments to disk, not these.
    downloadable: bool = True


class EmailSummary(ExchangeModel):
    id: str
    subject: str
    from_: EmailAddress = Field(alias="from")
    to: list[EmailAddress] = Field(default_factory=list)
    date: datetime
    is_read: bool
    has_attachments: bool = False
    preview: str = ""
    importance: Literal["low", "normal", "high"] = "normal"
    categories: list[str] = Field(default_factory=list)
    #: On the summary rather than EmailFull, so a listing can lead straight to
    #: get_thread without fetching every message first.
    conversation_id: str | None = None


class EmailFull(EmailSummary):
    cc: list[EmailAddress] = Field(default_factory=list)
    bcc: list[EmailAddress] = Field(default_factory=list)
    body_text: str = ""
    body_html: str | None = None
    attachments: list[Attachment] = Field(default_factory=list)
    headers: dict[str, str] = Field(default_factory=dict)
    truncated: bool = False


class Thread(ExchangeModel):
    #: Absent when the thread had to be rebuilt from its subject line.
    conversation_id: str | None = None
    subject: str
    message_count: int
    #: Oldest first.
    messages: list[EmailFull] = Field(default_factory=list)
    #: More messages exist than ``limit``; the oldest were dropped.
    truncated: bool = False


class Attendee(ExchangeModel):
    email: ServerAddress
    name: str | None = None
    response_type: Literal["accept", "tentative", "decline", "unknown"] = "unknown"


class CalendarEvent(ExchangeModel):
    """An appointment as the server holds it, not as we would have created it.

    No range check here on purpose. Exchange accepts zero-length appointments,
    and placeholders created by external systems arrive that way routinely;
    reversed timestamps show up in migrated mailboxes. list_events converts a
    whole page in one pass, so refusing a single such item used to fail every
    event around it — the user lost the day, not the bad appointment. Reporting
    it back is also what lets them find and fix it. The range checks that
    matter live on the request models below, where a mistake is still ours.
    """

    id: str
    subject: str
    start: datetime
    end: datetime
    location: str | None = None
    organizer: EmailAddress
    attendees: list[Attendee] = Field(default_factory=list)
    is_all_day: bool = False
    is_recurring: bool = False
    my_response: Literal["accept", "tentative", "decline", "unknown"] = "unknown"
    online_meeting_url: str | None = None
    body: str | None = None
    reminder_minutes: int | None = None
    categories: list[str] = Field(default_factory=list)
    recurrence_pattern: dict[str, Any] | None = None
    importance: Literal["low", "normal", "high"] = "normal"
    free_busy_status: Literal[
        "free", "tentative", "busy", "oof", "working_elsewhere", "unknown"
    ] = "busy"


class RecurrencePattern(ExchangeModel):
    type: Literal["daily", "weekly", "monthly", "yearly"]
    interval: int = Field(ge=1, default=1)
    end_date: date | None = None
    occurrences: int | None = Field(default=None, ge=1)
    days_of_week: list[WeekdayName] | None = None

    @model_validator(mode="after")
    def _validate_combination(self) -> "RecurrencePattern":
        if self.end_date is not None and self.occurrences is not None:
            raise ValueError("set either end_date or occurrences, not both")
        if self.days_of_week is not None and self.type != "weekly":
            raise ValueError("days_of_week only applies to type='weekly'")
        if self.type == "yearly" and self.interval != 1:
            # exchangelib's AbsoluteYearlyPattern has no interval field -- Exchange yearly
            # recurrence always repeats every year, so any other interval can't be honored.
            raise ValueError("interval is not supported for type='yearly'")
        return self


class ListEmailsRequest(ExchangeModel):
    folder: str = "inbox"
    limit: int = Field(default=20, ge=1, le=100)
    offset: int = Field(default=0, ge=0)
    from_address: EmailStr | None = None
    subject: str | None = None
    since: date | None = None
    before: date | None = None
    unread_only: bool = False
    has_attachments: bool | None = None


class GetEmailRequest(ExchangeModel):
    id: str


class GetThreadRequest(ExchangeModel):
    id: str | None = None
    conversation_id: str | None = None
    #: Sent is included by default: replies written from this mailbox live there,
    #: and a thread missing our own half of it is not useful.
    folders: list[str] = Field(default_factory=lambda: ["inbox", "sent"], min_length=1)
    limit: int = Field(default=20, ge=1, le=100)

    @model_validator(mode="after")
    def validate_selector(self) -> "GetThreadRequest":
        if bool(self.id) == bool(self.conversation_id):
            raise ValueError("exactly one of id or conversation_id must be provided")
        return self


class SendEmailRequest(ExchangeModel):
    to: list[EmailStr] = Field(min_length=1)
    subject: str = Field(min_length=1)
    body: str
    body_type: Literal["text", "html"] = "text"
    cc: list[EmailStr] = Field(default_factory=list)
    bcc: list[EmailStr] = Field(default_factory=list)
    reply_to: EmailStr | None = None
    attachments: list[Path] = Field(default_factory=list)
    importance: Literal["low", "normal", "high"] = "normal"
    include_signature: bool = True


class ReplyEmailRequest(ExchangeModel):
    id: str
    body: str
    reply_all: bool = False
    attachments: list[Path] = Field(default_factory=list)
    include_signature: bool = True


class ForwardEmailRequest(ExchangeModel):
    id: str
    to: list[EmailStr] = Field(min_length=1)
    comment: str | None = None
    attachments: list[Path] = Field(default_factory=list)
    include_signature: bool = True


class FolderActionRequest(ExchangeModel):
    id: str
    folder: str = Field(min_length=1)


class DeleteEmailRequest(ExchangeModel):
    id: str
    hard_delete: bool = False


class MarkEmailRequest(ExchangeModel):
    id: str
    read: bool | None = None
    flag: Literal["flagged", "complete", "none"] | None = None
    importance: Literal["low", "normal", "high"] | None = None
    flag_start_date: datetime | None = None
    flag_due_date: datetime | None = None

    @model_validator(mode="after")
    def validate_flag_dates(self) -> "MarkEmailRequest":
        if (
            self.flag_start_date is not None
            and self.flag_due_date is not None
            and self.flag_due_date < self.flag_start_date
        ):
            raise ValueError("flag_due_date must not be earlier than flag_start_date")
        if self.flag == "none" and (
            self.flag_start_date is not None or self.flag_due_date is not None
        ):
            raise ValueError("flag dates cannot be combined with flag='none'")
        return self


class InboxRuleConditions(ExchangeModel):
    #: Lenient on purpose: the server reports whatever a rule was created with,
    #: including X.500 distinguished names for internal senders. Read paths must
    #: survive that; outgoing rules use CreateInboxRuleConditions below.
    from_addresses: list[ServerAddress] = Field(default_factory=list)
    contains_sender_strings: list[str] = Field(default_factory=list)
    contains_subject_strings: list[str] = Field(default_factory=list)
    contains_subject_or_body_strings: list[str] = Field(default_factory=list)
    has_attachments: bool | None = None
    importance: Literal["low", "normal", "high"] | None = None

    def is_empty(self) -> bool:
        return not (
            self.from_addresses
            or self.contains_sender_strings
            or self.contains_subject_strings
            or self.contains_subject_or_body_strings
            or self.has_attachments is not None
            or self.importance is not None
        )


class InboxRuleActions(ExchangeModel):
    #: Folder name, path or id, resolved the same way move_email resolves it.
    move_to_folder: str | None = None
    assign_categories: list[str] = Field(default_factory=list)
    mark_as_read: bool | None = None
    delete: bool | None = None
    forward_to: list[ServerAddress] = Field(default_factory=list)

    def is_empty(self) -> bool:
        return not (
            self.move_to_folder
            or self.assign_categories
            or self.mark_as_read is not None
            or self.delete is not None
            or self.forward_to
        )


class CreateInboxRuleConditions(InboxRuleConditions):
    #: Strict where the base is lenient: a new rule must not be created against
    #: a garbage address, only read back with one.
    from_addresses: list[EmailStr] = Field(default_factory=list)


class CreateInboxRuleActions(InboxRuleActions):
    forward_to: list[EmailStr] = Field(default_factory=list)


class InboxRule(ExchangeModel):
    id: str | None = None
    display_name: str
    priority: int = 1
    is_enabled: bool = True
    #: Set by the server for rules it cannot fully express over EWS.
    is_not_supported: bool = False
    conditions: InboxRuleConditions = Field(default_factory=InboxRuleConditions)
    actions: InboxRuleActions = Field(default_factory=InboxRuleActions)


class CreateInboxRuleRequest(ExchangeModel):
    display_name: str = Field(min_length=1)
    priority: int = Field(default=1, ge=1)
    is_enabled: bool = True
    conditions: CreateInboxRuleConditions
    actions: CreateInboxRuleActions
    #: EWS documented behaviour: managing rules over EWS removes the client-side
    #: rule blob that desktop Outlook keeps, which can wipe rules created there.
    remove_outlook_rule_blob: bool = True

    @model_validator(mode="after")
    def validate_rule(self) -> "CreateInboxRuleRequest":
        if self.conditions.is_empty():
            raise ValueError("at least one condition is required")
        if self.actions.is_empty():
            raise ValueError("at least one action is required")
        return self


class UpdateInboxRuleRequest(ExchangeModel):
    id: str
    #: Only these two are updatable: a full update would have to round-trip the
    #: rule through this API's curated subset and silently drop any condition or
    #: action the subset does not model.
    is_enabled: bool | None = None
    priority: int | None = Field(default=None, ge=1)
    remove_outlook_rule_blob: bool = True

    @model_validator(mode="after")
    def validate_update(self) -> "UpdateInboxRuleRequest":
        if self.is_enabled is None and self.priority is None:
            raise ValueError("nothing to update: set is_enabled and/or priority")
        return self


class DeleteInboxRuleRequest(ExchangeModel):
    id: str
    remove_outlook_rule_blob: bool = True


class DelegatePermissionLevels(ExchangeModel):
    calendar: str = "None"
    inbox: str = "None"
    tasks: str = "None"
    contacts: str = "None"


class DelegateInfo(ExchangeModel):
    email: ServerAddress | None = None
    display_name: str | None = None
    permissions: DelegatePermissionLevels = Field(default_factory=DelegatePermissionLevels)
    receives_copies_of_meeting_messages: bool = False
    can_view_private_items: bool = False


class OutOfOfficeSettings(ExchangeModel):
    state: Literal["disabled", "enabled", "scheduled"]
    #: Who outside the organisation gets the external reply.
    external_audience: Literal["none", "known", "all"] = "all"
    internal_reply: str | None = None
    external_reply: str | None = None
    #: Only meaningful (and required) when state is "scheduled".
    start: datetime | None = None
    end: datetime | None = None

    @model_validator(mode="after")
    def validate_schedule(self) -> "OutOfOfficeSettings":
        if self.state == "scheduled":
            if self.start is None or self.end is None:
                raise ValueError("state='scheduled' requires both start and end")
            if self.end <= self.start:
                raise ValueError("end must be greater than start")
        return self


class AddAttachmentRequest(ExchangeModel):
    email_id: str
    #: Local file to attach; must live under EXCHANGE_ATTACHMENT_ROOT.
    path: Path


class DeleteAttachmentRequest(ExchangeModel):
    email_id: str
    attachment_id: str


class BulkMoveEmailsRequest(ExchangeModel):
    ids: list[str] = Field(min_length=1, max_length=500)
    folder: str = Field(min_length=1)


class BulkDeleteEmailsRequest(ExchangeModel):
    ids: list[str] = Field(min_length=1, max_length=500)
    hard_delete: bool = False


class BulkItemResult(ExchangeModel):
    id: str
    #: Move/copy give the item a new id; the old one is dead after the operation.
    new_id: str | None = None


class BulkItemFailure(ExchangeModel):
    id: str
    error: str
    message: str


class BulkResult(ExchangeModel):
    succeeded: list[BulkItemResult] = Field(default_factory=list)
    failed: list[BulkItemFailure] = Field(default_factory=list)


class CategorizeEmailRequest(ExchangeModel):
    id: str
    categories: list[str] = Field(default_factory=list)
    #: ``set`` with an empty list clears every category on the message.
    mode: Literal["set", "add", "remove"] = "set"

    @model_validator(mode="after")
    def validate_categories(self) -> "CategorizeEmailRequest":
        if self.mode in {"add", "remove"} and not self.categories:
            raise ValueError(f"categories must not be empty when mode is '{self.mode}'")
        if any(not name.strip() for name in self.categories):
            raise ValueError("category names must not be blank")
        return self


class ListCategoriesRequest(ExchangeModel):
    folders: list[str] = Field(default_factory=lambda: ["inbox"], min_length=1)
    limit: int = Field(default=200, ge=1, le=1000)


class CategoryUsage(ExchangeModel):
    name: str
    count: int
    #: Outlook preset colour name; None when the mailbox master list is
    #: unavailable and counts were collected from recent messages instead.
    color: str | None = None


class ListFoldersRequest(ExchangeModel):
    parent: str | None = None
    depth: int = Field(default=2, ge=0, le=10)


class CreateFolderRequest(ExchangeModel):
    name: str = Field(min_length=1)
    parent: str | None = "inbox"


class RenameFolderRequest(ExchangeModel):
    folder: str = Field(min_length=1)
    name: str = Field(min_length=1)


class DeleteFolderRequest(ExchangeModel):
    folder: str = Field(min_length=1)
    hard_delete: bool = False


class DraftEmailRequest(ExchangeModel):
    to: list[EmailStr] = Field(min_length=1)
    subject: str = Field(min_length=1)
    body: str
    body_type: Literal["text", "html"] = "text"
    cc: list[EmailStr] = Field(default_factory=list)
    bcc: list[EmailStr] = Field(default_factory=list)
    attachments: list[Path] = Field(default_factory=list)
    include_signature: bool = True


class SendDraftRequest(ExchangeModel):
    id: str


class SearchEmailsRequest(ExchangeModel):
    #: Substring match over subject, body and sender.
    query: str | None = Field(default=None, min_length=1)
    #: EWS Advanced Query Syntax, e.g. 'from:ivan AND hasattachments:true'.
    #: Server-side full-text search; needs content indexing enabled on the server.
    aqs: str | None = Field(default=None, min_length=1)
    folder: str | None = None
    limit: int = Field(default=20, ge=1, le=100)

    @model_validator(mode="after")
    def validate_selector(self) -> "SearchEmailsRequest":
        if bool(self.query) == bool(self.aqs):
            raise ValueError("exactly one of query or aqs must be provided")
        return self


class GetAttachmentRequest(ExchangeModel):
    email_id: str
    attachment_id: str
    save_path: Path | None = None


class ListEventsRequest(ExchangeModel):
    start: datetime
    end: datetime
    calendar_id: str | None = None
    include_recurring: bool = True

    @model_validator(mode="after")
    def validate_range(self) -> "ListEventsRequest":
        if self.end <= self.start:
            raise ValueError("end must be greater than start")
        return self


class GetEventRequest(ExchangeModel):
    id: str
    calendar_id: str | None = None


class WorkHours(ExchangeModel):
    start: str = "09:00"
    end: str = "18:00"

    @field_validator("start", "end")
    @classmethod
    def _validate_format(cls, value: str) -> str:
        if not _HHMM_RE.match(value):
            raise ValueError("must be 24-hour HH:MM, e.g. '09:00'")
        return value

    @model_validator(mode="after")
    def _validate_order(self) -> "WorkHours":
        if self.end <= self.start:
            raise ValueError("end must be after start")
        return self


class CreateEventRequest(ExchangeModel):
    subject: str = Field(min_length=1)
    start: datetime
    end: datetime
    calendar_id: str | None = None
    location: str | None = None
    body: str | None = None
    attendees: list[EmailStr] = Field(default_factory=list)
    is_all_day: bool = False
    reminder_minutes: int | None = Field(default=15, ge=0, le=10080)
    recurrence: RecurrencePattern | None = None
    categories: list[str] = Field(default_factory=list)
    importance: Literal["low", "normal", "high"] = "normal"

    @model_validator(mode="after")
    def validate_range(self) -> "CreateEventRequest":
        if self.end <= self.start:
            raise ValueError("end must be greater than start")
        if (
            self.recurrence
            and self.recurrence.end_date is not None
            and self.recurrence.end_date < self.start.date()
        ):
            raise ValueError("recurrence.end_date must not be before start")
        return self


class UpdateEventRequest(ExchangeModel):
    id: str
    calendar_id: str | None = None
    subject: str | None = None
    start: datetime | None = None
    end: datetime | None = None
    location: str | None = None
    body: str | None = None
    add_attendees: list[EmailStr] = Field(default_factory=list)
    remove_attendees: list[EmailStr] = Field(default_factory=list)
    reminder_minutes: int | None = Field(default=None, ge=0, le=10080)
    send_updates: Literal["none", "all", "modified"] = "all"

    @model_validator(mode="after")
    def validate_range(self) -> "UpdateEventRequest":
        if self.start and self.end and self.end <= self.start:
            raise ValueError("end must be greater than start")
        return self


class DeleteEventRequest(ExchangeModel):
    id: str
    calendar_id: str | None = None
    notify_attendees: bool = True
    cancel_message: str | None = None


class RespondToInviteRequest(ExchangeModel):
    id: str
    calendar_id: str | None = None
    response: Literal["accept", "tentative", "decline"]
    message: str | None = None


class FindFreeSlotsRequest(ExchangeModel):
    attendees: list[EmailStr] = Field(min_length=1)
    duration: int = Field(ge=1, le=1440)
    start: datetime
    end: datetime
    work_hours: WorkHours | None = None

    @model_validator(mode="after")
    def validate_range(self) -> "FindFreeSlotsRequest":
        if self.end <= self.start:
            raise ValueError("end must be greater than start")
        return self


class FreeSlot(ExchangeModel):
    start: datetime
    end: datetime
    all_available: bool = True
    busy_attendees: list[ServerAddress] = Field(default_factory=list)

    @model_validator(mode="after")
    def validate_range(self) -> "FreeSlot":
        if self.end <= self.start:
            raise ValueError("end must be greater than start")
        return self


class SendResult(ExchangeModel):
    # EWS doesn't hand back a usable id for a send: send-and-save typically returns
    # none at all, and a reply/forward/draft-send's only candidate ids (the source
    # item, the now-invalidated draft) refer to a different or no-longer-valid item.
    id: str | None = None
    status: str
    warning: str | None = None


class CreateEventResult(ExchangeModel):
    id: str
    status: str
    subject: str
    start: datetime
    end: datetime
    invite_sent: bool
    warning: str | None = None


class PingResult(ExchangeModel):
    status: Literal["ok"]
    server: str
    version: str | None = None
    latency_ms: int | None = None


class MailboxInfo(ExchangeModel):
    email_address: ServerAddress
    display_name: str
    timezone: str
    mailbox_size_mb: float | None = None
    quota_mb: float | None = None
    exchange_version: str | None = None


class FolderInfo(ExchangeModel):
    id: str | None = None
    name: str
    path: str
    unread_count: int = 0
    total_count: int = 0
    children: list["FolderInfo"] = Field(default_factory=list)


class ActionResult(ExchangeModel):
    id: str
    status: str
    updated_fields: list[str] = Field(default_factory=list)
    warning: str | None = None
    new_folder: str | None = None
    new_id: str | None = None
    path: str | None = None
    #: Categories the message carries after the change.
    categories: list[str] | None = None


class AttachmentResult(ExchangeModel):
    filename: str
    size: int
    saved_path: str
    content_type: str | None = None


class CalendarInfo(ExchangeModel):
    id: str
    name: str
    is_default: bool = False
    color: str | None = None
    owner_email: ServerAddress | None = None


class AvailabilityResult(ExchangeModel):
    free_slots: list[FreeSlot] = Field(default_factory=list)
    busy_slots: list[dict[str, Any]] = Field(default_factory=list)


class ContactSummary(ExchangeModel):
    id: str
    display_name: str
    email_addresses: list[str] = Field(default_factory=list)
    phone_numbers: list[str] = Field(default_factory=list)
    company: str | None = None
    job_title: str | None = None
    department: str | None = None
    source: Literal["personal", "gal"]


class ContactEmailAddress(ExchangeModel):
    type: str
    address: ServerAddress


class ContactPhoneNumber(ExchangeModel):
    type: str
    number: str


class ContactAddress(ExchangeModel):
    type: str
    street: str | None = None
    city: str | None = None
    state: str | None = None
    postal_code: str | None = None
    country: str | None = None


class ContactFull(ExchangeModel):
    id: str
    display_name: str
    first_name: str | None = None
    last_name: str | None = None
    email_addresses: list[ContactEmailAddress] = Field(default_factory=list)
    phone_numbers: list[ContactPhoneNumber] = Field(default_factory=list)
    addresses: list[ContactAddress] = Field(default_factory=list)
    company: str | None = None
    job_title: str | None = None
    department: str | None = None
    manager: str | None = None
    notes: str | None = None
    photo_url: str | None = None
    birthday: date | None = None
    source: Literal["personal", "gal"]


class SearchContactsRequest(ExchangeModel):
    query: str = Field(min_length=1)
    source: Literal["personal", "gal", "all"] = "all"
    limit: int = Field(default=10, ge=1, le=100)


class GetContactRequest(ExchangeModel):
    id: str
    source: Literal["personal", "gal"] | None = None


class CreateContactRequest(ExchangeModel):
    display_name: str = Field(min_length=1)
    first_name: str | None = None
    last_name: str | None = None
    email: EmailStr | None = None
    phone: str | None = None
    company: str | None = None
    job_title: str | None = None
    notes: str | None = None


class UpdateContactRequest(ExchangeModel):
    """A partial update: omitted fields are left unchanged; a field explicitly
    set to ``null`` clears that value on the contact (see ``model_fields_set``
    usage in ``exchange_client/contacts.py``'s ``update_contact``)."""

    id: str
    display_name: str | None = None
    first_name: str | None = None
    last_name: str | None = None
    email: EmailStr | None = None
    phone: str | None = None
    company: str | None = None
    job_title: str | None = None
    notes: str | None = None


class DeleteContactRequest(ExchangeModel):
    id: str
    hard_delete: bool = False


def dump_model(
    value: BaseModel | Sequence[BaseModel] | dict[str, Any] | Sequence[dict[str, Any]],
) -> Any:
    if isinstance(value, BaseModel):
        return value.model_dump(mode="json", by_alias=True)
    if isinstance(value, list):
        return [dump_model(item) for item in value]
    if isinstance(value, dict):
        return value
    return value
