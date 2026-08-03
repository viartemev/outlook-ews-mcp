from __future__ import annotations

from collections.abc import Sequence
from datetime import date, datetime, time
from pathlib import Path
from typing import Annotated, Any, Literal

from pydantic import BaseModel, BeforeValidator, ConfigDict, EmailStr, Field, model_validator


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


class EmailFull(EmailSummary):
    cc: list[EmailAddress] = Field(default_factory=list)
    bcc: list[EmailAddress] = Field(default_factory=list)
    body_text: str = ""
    body_html: str | None = None
    attachments: list[Attachment] = Field(default_factory=list)
    conversation_id: str | None = None
    headers: dict[str, str] = Field(default_factory=dict)
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


class RecurrencePattern(ExchangeModel):
    type: Literal["daily", "weekly", "monthly", "yearly"]
    interval: int = Field(ge=1, default=1)
    end_date: date | None = None
    occurrences: int | None = Field(default=None, ge=1)
    days_of_week: list[str] | None = None


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


class ListFoldersRequest(ExchangeModel):
    parent: str | None = None
    depth: int = Field(default=2, ge=0, le=10)


class CreateFolderRequest(ExchangeModel):
    name: str = Field(min_length=1)
    parent: str | None = "inbox"


class DraftEmailRequest(ExchangeModel):
    to: list[EmailStr] = Field(min_length=1)
    subject: str = Field(min_length=1)
    body: str
    body_type: Literal["text", "html"] = "text"
    cc: list[EmailStr] = Field(default_factory=list)
    bcc: list[EmailStr] = Field(default_factory=list)
    attachments: list[Path] = Field(default_factory=list)


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


class WorkHours(ExchangeModel):
    start: str = "09:00"
    end: str = "18:00"


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
        if self.is_all_day:
            self.start = datetime.combine(self.start.date(), time.min, tzinfo=self.start.tzinfo)
            self.end = datetime.combine(self.end.date(), time.max, tzinfo=self.end.tzinfo)
        return self


class UpdateEventRequest(ExchangeModel):
    id: str
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
    notify_attendees: bool = True
    cancel_message: str | None = None


class RespondToInviteRequest(ExchangeModel):
    id: str
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
    id: str
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


class CreateContactRequest(ExchangeModel):
    display_name: str = Field(min_length=1)
    first_name: str | None = None
    last_name: str | None = None
    email: EmailStr | None = None
    phone: str | None = None
    company: str | None = None
    job_title: str | None = None
    notes: str | None = None


class UpdateContactRequest(CreateContactRequest):
    id: str


class DeleteContactRequest(ExchangeModel):
    id: str


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
