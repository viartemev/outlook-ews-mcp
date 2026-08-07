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


class ReplyEmailRequest(ExchangeModel):
    id: str
    body: str
    reply_all: bool = False
    attachments: list[Path] = Field(default_factory=list)


class ForwardEmailRequest(ExchangeModel):
    id: str
    to: list[EmailStr] = Field(min_length=1)
    comment: str | None = None
    attachments: list[Path] = Field(default_factory=list)


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


class BulkMoveEmailsRequest(ExchangeModel):
    ids: list[str] = Field(min_length=1, max_length=50)
    folder: str = Field(min_length=1)


class BulkDeleteEmailsRequest(ExchangeModel):
    ids: list[str] = Field(min_length=1, max_length=50)
    hard_delete: bool = False


class BulkMarkEmailsRequest(ExchangeModel):
    ids: list[str] = Field(min_length=1, max_length=50)
    read: bool | None = None
    flag: Literal["flagged", "complete", "none"] | None = None
    importance: Literal["low", "normal", "high"] | None = None
    flag_start_date: datetime | None = None
    flag_due_date: datetime | None = None

    @model_validator(mode="after")
    def validate_flag_dates(self) -> "BulkMarkEmailsRequest":
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


class BulkCategorizeEmailsRequest(ExchangeModel):
    ids: list[str] = Field(min_length=1, max_length=50)
    categories: list[str] = Field(default_factory=list)
    mode: Literal["set", "add", "remove"] = "set"

    @model_validator(mode="after")
    def validate_categories(self) -> "BulkCategorizeEmailsRequest":
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


class UpdateDraftRequest(ExchangeModel):
    """A partial update: omitted fields are left unchanged. ``attachments``, when
    present, replaces the draft's entire attachment set (there is no add/remove
    mode -- see ``model_fields_set`` usage in ``exchange_client/email.py``'s
    ``update_draft``)."""

    id: str
    to: list[EmailStr] | None = None
    subject: str | None = Field(default=None, min_length=1)
    body: str | None = None
    body_type: Literal["text", "html"] = "text"
    cc: list[EmailStr] | None = None
    bcc: list[EmailStr] | None = None
    attachments: list[Path] | None = None


class SendDraftRequest(ExchangeModel):
    id: str


class SearchEmailsRequest(ExchangeModel):
    query: str = Field(min_length=1)
    folder: str | None = None
    limit: int = Field(default=20, ge=1, le=100)


class GetAttachmentRequest(ExchangeModel):
    email_id: str
    attachment_id: str
    save_path: Path | None = None


class GetEmailMimeRequest(ExchangeModel):
    id: str


class EmailMimeResult(ExchangeModel):
    id: str
    filename: str
    content_type: str = "message/rfc822"
    size: int
    #: Raw RFC 822 message, base64-encoded (the source can contain arbitrary
    #: bytes -- inline attachment payloads, non-UTF-8 header encodings -- so
    #: it isn't safe to hand back as plain text).
    mime_base64: str


def _validate_rule_has_action(rule: Any) -> None:
    if not any(
        [
            rule.move_to_folder is not None,
            rule.mark_as_read is not None,
            rule.assign_categories,
            rule.delete,
        ]
    ):
        raise ValueError(
            "at least one action (move_to_folder, mark_as_read, assign_categories, "
            "delete) must be set"
        )


class MailRule(ExchangeModel):
    """An Inbox rule as the server holds it.

    Only a focused subset of EWS's condition/action vocabulary is exposed --
    the fields most tool callers actually need (sender/subject/attachment
    conditions; move/mark/categorize/delete actions) -- rather than the full
    Conditions/Actions surface exchangelib exposes.
    """

    id: str
    display_name: str
    priority: int
    is_enabled: bool = True
    from_addresses: list[str] = Field(default_factory=list)
    contains_subject_strings: list[str] = Field(default_factory=list)
    has_attachments: bool | None = None
    move_to_folder: str | None = None
    mark_as_read: bool | None = None
    assign_categories: list[str] = Field(default_factory=list)
    delete: bool = False
    stop_processing_rules: bool = True


class CreateRuleRequest(ExchangeModel):
    display_name: str = Field(min_length=1)
    priority: int = Field(default=1, ge=1)
    is_enabled: bool = True
    #: Conditions -- a rule with none of these set matches every message.
    from_addresses: list[EmailStr] = Field(default_factory=list)
    contains_subject_strings: list[str] = Field(default_factory=list)
    has_attachments: bool | None = None
    #: Actions -- at least one must be set (see validate_has_action below).
    move_to_folder: str | None = None
    mark_as_read: bool | None = None
    assign_categories: list[str] = Field(default_factory=list)
    delete: bool = False
    stop_processing_rules: bool = True

    @model_validator(mode="after")
    def validate_has_action(self) -> "CreateRuleRequest":
        _validate_rule_has_action(self)
        return self


class UpdateRuleRequest(ExchangeModel):
    """A full replace, not a partial patch: EWS's SetInboxRule replaces the whole
    rule server-side, so every field here -- not just the ones being changed --
    must reflect the rule's desired end state."""

    id: str = Field(min_length=1)
    display_name: str = Field(min_length=1)
    priority: int = Field(default=1, ge=1)
    is_enabled: bool = True
    from_addresses: list[EmailStr] = Field(default_factory=list)
    contains_subject_strings: list[str] = Field(default_factory=list)
    has_attachments: bool | None = None
    move_to_folder: str | None = None
    mark_as_read: bool | None = None
    assign_categories: list[str] = Field(default_factory=list)
    delete: bool = False
    stop_processing_rules: bool = True

    @model_validator(mode="after")
    def validate_has_action(self) -> "UpdateRuleRequest":
        _validate_rule_has_action(self)
        return self


class DeleteRuleRequest(ExchangeModel):
    id: str = Field(min_length=1)


OofState = Literal["disabled", "enabled", "scheduled"]
OofExternalAudience = Literal["none", "known", "all"]


class OofSettingsModel(ExchangeModel):
    state: OofState
    external_audience: OofExternalAudience = "all"
    start: datetime | None = None
    end: datetime | None = None
    internal_reply: str | None = None
    external_reply: str | None = None


class SetOofSettingsRequest(ExchangeModel):
    state: OofState
    external_audience: OofExternalAudience = "all"
    start: datetime | None = None
    end: datetime | None = None
    internal_reply: str | None = None
    external_reply: str | None = None

    @model_validator(mode="after")
    def validate_state(self) -> "SetOofSettingsRequest":
        # Mirrors the validation exchangelib's OofSettings.clean() applies
        # server-side, surfaced here instead of as an opaque Exchange error.
        if self.state == "scheduled":
            if self.start is None or self.end is None:
                raise ValueError("start and end are required when state='scheduled'")
            if self.end <= self.start:
                raise ValueError("end must be greater than start")
        if self.state != "disabled" and (not self.internal_reply or not self.external_reply):
            raise ValueError(
                "internal_reply and external_reply are required unless state='disabled'"
            )
        return self


class ListEventsRequest(ExchangeModel):
    start: datetime
    end: datetime
    calendar_id: str | None = None
    include_recurring: bool = True
    #: View another mailbox's default calendar instead of the service account's own
    #: (requires delegate/impersonation access to that mailbox on the server side).
    #: Not combinable with calendar_id -- mailbox scoping only ever targets that
    #: mailbox's default calendar.
    mailbox: EmailStr | None = None

    @model_validator(mode="after")
    def validate_range(self) -> "ListEventsRequest":
        if self.end <= self.start:
            raise ValueError("end must be greater than start")
        if self.mailbox is not None and self.calendar_id is not None:
            raise ValueError("mailbox cannot be combined with calendar_id")
        return self


class GetEventRequest(ExchangeModel):
    id: str
    calendar_id: str | None = None
    mailbox: EmailStr | None = None

    @model_validator(mode="after")
    def validate_mailbox(self) -> "GetEventRequest":
        if self.mailbox is not None and self.calendar_id is not None:
            raise ValueError("mailbox cannot be combined with calendar_id")
        return self


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


class RoomListInfo(ExchangeModel):
    name: str
    email: ServerAddress


class RoomInfo(ExchangeModel):
    name: str
    email: ServerAddress


class ListRoomsRequest(ExchangeModel):
    room_list: EmailStr


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
