from __future__ import annotations

import logging
import re
import tempfile
import warnings
from collections.abc import Iterable
from datetime import UTC, datetime, timedelta
from pathlib import Path
from typing import Any, Protocol
from urllib.parse import urlparse

from exchangelib import (
    Account,
    Attendee,
    BASIC,
    CalendarItem,
    Configuration,
    Credentials,
    DELEGATE,
    EWSTimeZone,
    FileAttachment,
    Folder,
    HTMLBody,
    IMPERSONATION,
    Mailbox,
    Message,
    NTLM,
)
from exchangelib.errors import (
    ErrorItemSavePropertyError,
    ErrorFolderSavePropertyError,
    RateLimitError,
    ResponseMessageError,
    TransportError,
    UnknownTimeZone,
    UnauthorizedError,
)
from exchangelib.indexed_properties import EmailAddress as IndexedEmailAddress
from exchangelib.indexed_properties import PhoneNumber as IndexedPhoneNumber
from exchangelib.items import Contact
from exchangelib.properties import FreeBusyViewOptions, ItemId, MailboxData, TimeWindow
from exchangelib.protocol import BaseProtocol, FailFast, FaultTolerance
from exchangelib.services import GetUserAvailability, ResolveNames
from urllib3.exceptions import InsecureRequestWarning

from .auth import build_auth_context
from .config import Settings
from .errors import (
    APIError,
    AuthFailedError,
    ConflictError,
    ExchangeUnavailableError,
    NotFoundError,
    PermissionDeniedError,
    TimeoutAPIError,
)
from .models import (
    ActionResult,
    AttachmentResult,
    AvailabilityResult,
    CalendarInfo,
    CalendarEvent,
    ContactFull,
    ContactSummary,
    CreateEventRequest,
    CreateEventResult,
    CreateContactRequest,
    CreateFolderRequest,
    DeleteContactRequest,
    DeleteEmailRequest,
    DeleteEventRequest,
    DraftEmailRequest,
    EmailFull,
    EmailSummary,
    EmailAddress,
    FolderActionRequest,
    FolderInfo,
    FindFreeSlotsRequest,
    ForwardEmailRequest,
    FreeSlot,
    GetAttachmentRequest,
    GetContactRequest,
    GetEmailRequest,
    GetEventRequest,
    ListEmailsRequest,
    ListEventsRequest,
    ListFoldersRequest,
    MailboxInfo,
    MarkEmailRequest,
    PingResult,
    ReplyEmailRequest,
    RespondToInviteRequest,
    SearchContactsRequest,
    SearchEmailsRequest,
    SendDraftRequest,
    SendEmailRequest,
    SendResult,
    Attachment,
    UpdateContactRequest,
    UpdateEventRequest,
)

logger = logging.getLogger(__name__)
_RAW_SESSION_PATCHED = False
_TIMEZONE_FALLBACK_PATCHED = False
_GUID_TIMEZONE_RE = re.compile(
    r"^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$"
)


class ExchangeBackend(Protocol):
    def ping(self) -> PingResult: ...
    def get_mailbox_info(self) -> MailboxInfo: ...
    def list_emails(self, request: ListEmailsRequest) -> list[EmailSummary]: ...
    def get_email(self, request: GetEmailRequest) -> EmailFull: ...
    def search_emails(self, request: SearchEmailsRequest) -> list[EmailSummary]: ...
    def send_email(self, request: SendEmailRequest) -> SendResult: ...
    def reply_email(self, request: ReplyEmailRequest) -> SendResult: ...
    def forward_email(self, request: ForwardEmailRequest) -> SendResult: ...
    def move_email(self, request: FolderActionRequest) -> ActionResult: ...
    def copy_email(self, request: FolderActionRequest) -> ActionResult: ...
    def delete_email(self, request: DeleteEmailRequest) -> ActionResult: ...
    def mark_email(self, request: MarkEmailRequest) -> ActionResult: ...
    def list_folders(self, request: ListFoldersRequest) -> list[FolderInfo]: ...
    def create_folder(self, request: CreateFolderRequest) -> ActionResult: ...
    def create_draft(self, request: DraftEmailRequest) -> ActionResult: ...
    def send_draft(self, request: SendDraftRequest) -> ActionResult: ...
    def get_attachment(self, request: GetAttachmentRequest) -> AttachmentResult: ...
    def list_events(self, request: ListEventsRequest) -> list[CalendarEvent]: ...
    def get_event(self, request: GetEventRequest) -> CalendarEvent: ...
    def create_event(self, request: CreateEventRequest) -> CreateEventResult: ...
    def update_event(self, request: UpdateEventRequest) -> ActionResult: ...
    def delete_event(self, request: DeleteEventRequest) -> ActionResult: ...
    def respond_to_invite(self, request: RespondToInviteRequest) -> ActionResult: ...
    def find_free_slots(self, request: FindFreeSlotsRequest) -> list[FreeSlot]: ...
    def get_my_availability(self, request: ListEventsRequest) -> AvailabilityResult: ...
    def list_calendars(self) -> list[CalendarInfo]: ...
    def search_contacts(self, request: SearchContactsRequest) -> list[ContactSummary]: ...
    def get_contact(self, request: GetContactRequest) -> ContactFull: ...
    def create_contact(self, request: CreateContactRequest) -> ActionResult: ...
    def update_contact(self, request: UpdateContactRequest) -> ActionResult: ...
    def delete_contact(self, request: DeleteContactRequest) -> ActionResult: ...


class UnconfiguredExchangeBackend:
    def __init__(self, settings: Settings) -> None:
        self.settings = settings

    def _raise(self) -> None:
        raise ExchangeUnavailableError(
            "exchange backend is not configured; provide a real EWS-backed implementation"
        )

    def ping(self) -> PingResult:
        self._raise()

    def get_mailbox_info(self) -> MailboxInfo:
        self._raise()

    def list_emails(self, request: ListEmailsRequest) -> list[EmailSummary]:
        self._raise()

    def get_email(self, request: GetEmailRequest) -> EmailFull:
        self._raise()

    def send_email(self, request: SendEmailRequest) -> SendResult:
        self._raise()

    def search_emails(self, request: SearchEmailsRequest) -> list[EmailSummary]:
        self._raise()

    def reply_email(self, request: ReplyEmailRequest) -> SendResult:
        self._raise()

    def forward_email(self, request: ForwardEmailRequest) -> SendResult:
        self._raise()

    def move_email(self, request: FolderActionRequest) -> ActionResult:
        self._raise()

    def copy_email(self, request: FolderActionRequest) -> ActionResult:
        self._raise()

    def delete_email(self, request: DeleteEmailRequest) -> ActionResult:
        self._raise()

    def mark_email(self, request: MarkEmailRequest) -> ActionResult:
        self._raise()

    def list_folders(self, request: ListFoldersRequest) -> list[FolderInfo]:
        self._raise()

    def create_folder(self, request: CreateFolderRequest) -> ActionResult:
        self._raise()

    def create_draft(self, request: DraftEmailRequest) -> ActionResult:
        self._raise()

    def send_draft(self, request: SendDraftRequest) -> ActionResult:
        self._raise()

    def get_attachment(self, request: GetAttachmentRequest) -> AttachmentResult:
        self._raise()

    def list_events(self, request: ListEventsRequest) -> list[CalendarEvent]:
        self._raise()

    def get_event(self, request: GetEventRequest) -> CalendarEvent:
        self._raise()

    def create_event(self, request: CreateEventRequest) -> CreateEventResult:
        self._raise()

    def update_event(self, request: UpdateEventRequest) -> ActionResult:
        self._raise()

    def delete_event(self, request: DeleteEventRequest) -> ActionResult:
        self._raise()

    def respond_to_invite(self, request: RespondToInviteRequest) -> ActionResult:
        self._raise()

    def find_free_slots(self, request: FindFreeSlotsRequest) -> list[FreeSlot]:
        self._raise()

    def get_my_availability(self, request: ListEventsRequest) -> AvailabilityResult:
        self._raise()

    def list_calendars(self) -> list[CalendarInfo]:
        self._raise()

    def search_contacts(self, request: SearchContactsRequest) -> list[ContactSummary]:
        self._raise()

    def get_contact(self, request: GetContactRequest) -> ContactFull:
        self._raise()

    def create_contact(self, request: CreateContactRequest) -> ActionResult:
        self._raise()

    def update_contact(self, request: UpdateContactRequest) -> ActionResult:
        self._raise()

    def delete_contact(self, request: DeleteContactRequest) -> ActionResult:
        self._raise()


class ExchangeClient:
    def __init__(self, settings: Settings, backend: ExchangeBackend | None = None) -> None:
        self.settings = settings
        self.backend = backend or UnconfiguredExchangeBackend(settings)

    def ping(self) -> PingResult:
        return self.backend.ping()

    def get_mailbox_info(self) -> MailboxInfo:
        return self.backend.get_mailbox_info()

    def list_emails(self, request: ListEmailsRequest) -> list[EmailSummary]:
        return self.backend.list_emails(request)

    def get_email(self, request: GetEmailRequest) -> EmailFull:
        return self.backend.get_email(request)

    def search_emails(self, request: SearchEmailsRequest) -> list[EmailSummary]:
        return self.backend.search_emails(request)

    def send_email(self, request: SendEmailRequest) -> SendResult:
        return self.backend.send_email(request)

    def reply_email(self, request: ReplyEmailRequest) -> SendResult:
        return self.backend.reply_email(request)

    def forward_email(self, request: ForwardEmailRequest) -> SendResult:
        return self.backend.forward_email(request)

    def move_email(self, request: FolderActionRequest) -> ActionResult:
        return self.backend.move_email(request)

    def copy_email(self, request: FolderActionRequest) -> ActionResult:
        return self.backend.copy_email(request)

    def delete_email(self, request: DeleteEmailRequest) -> ActionResult:
        return self.backend.delete_email(request)

    def mark_email(self, request: MarkEmailRequest) -> ActionResult:
        return self.backend.mark_email(request)

    def list_folders(self, request: ListFoldersRequest) -> list[FolderInfo]:
        return self.backend.list_folders(request)

    def create_folder(self, request: CreateFolderRequest) -> ActionResult:
        return self.backend.create_folder(request)

    def create_draft(self, request: DraftEmailRequest) -> ActionResult:
        return self.backend.create_draft(request)

    def send_draft(self, request: SendDraftRequest) -> ActionResult:
        return self.backend.send_draft(request)

    def get_attachment(self, request: GetAttachmentRequest) -> AttachmentResult:
        return self.backend.get_attachment(request)

    def list_events(self, request: ListEventsRequest) -> list[CalendarEvent]:
        return self.backend.list_events(request)

    def get_event(self, request: GetEventRequest) -> CalendarEvent:
        return self.backend.get_event(request)

    def create_event(self, request: CreateEventRequest) -> CreateEventResult:
        return self.backend.create_event(request)

    def update_event(self, request: UpdateEventRequest) -> ActionResult:
        return self.backend.update_event(request)

    def delete_event(self, request: DeleteEventRequest) -> ActionResult:
        return self.backend.delete_event(request)

    def respond_to_invite(self, request: RespondToInviteRequest) -> ActionResult:
        return self.backend.respond_to_invite(request)

    def find_free_slots(self, request: FindFreeSlotsRequest) -> list[FreeSlot]:
        return self.backend.find_free_slots(request)

    def get_my_availability(self, request: ListEventsRequest) -> AvailabilityResult:
        return self.backend.get_my_availability(request)

    def list_calendars(self) -> list[CalendarInfo]:
        return self.backend.list_calendars()

    def search_contacts(self, request: SearchContactsRequest) -> list[ContactSummary]:
        return self.backend.search_contacts(request)

    def get_contact(self, request: GetContactRequest) -> ContactFull:
        return self.backend.get_contact(request)

    def create_contact(self, request: CreateContactRequest) -> ActionResult:
        return self.backend.create_contact(request)

    def update_contact(self, request: UpdateContactRequest) -> ActionResult:
        return self.backend.update_contact(request)

    def delete_contact(self, request: DeleteContactRequest) -> ActionResult:
        return self.backend.delete_contact(request)


def build_default_backend(settings: Settings) -> ExchangeBackend:
    return EWSExchangeBackend(settings)


class EWSExchangeBackend:
    def __init__(self, settings: Settings) -> None:
        self.settings = settings
        self._account: Account | None = None

    @property
    def account(self) -> Account:
        if self._account is None:
            self._account = self._build_account()
        return self._account

    def _build_account(self) -> Account:
        auth = build_auth_context(self.settings)
        if auth.auth_type == "OAuth2":
            raise APIError(
                "validation_error",
                "OAuth2 is not wired in this build yet",
                details=[{"field": "EXCHANGE_AUTH_TYPE", "reason": "supported values for live checks are NTLM or Basic"}],
            )

        BaseProtocol.TIMEOUT = self.settings.exchange_timeout
        self._configure_ssl_verification()
        self._configure_timezone_fallback()
        retry_policy = FailFast() if self.settings.exchange_max_retries == 0 else FaultTolerance(
            max_wait=max(self.settings.exchange_timeout * self.settings.exchange_max_retries, self.settings.exchange_timeout)
        )
        credentials = Credentials(username=auth.username, password=auth.password)
        auth_type = BASIC if auth.auth_type == "Basic" else NTLM
        service_endpoint = self._normalize_service_endpoint(self.settings.exchange_server)
        config = Configuration(
            service_endpoint=service_endpoint,
            credentials=credentials,
            auth_type=auth_type,
            retry_policy=retry_policy,
        )
        access_type = IMPERSONATION if auth.impersonate_as else DELEGATE
        try:
            return Account(
                primary_smtp_address=auth.primary_smtp_address,
                config=config,
                autodiscover=False,
                access_type=access_type,
            )
        except Exception as exc:  # noqa: BLE001
            raise self._map_exception(exc) from exc

    def _normalize_service_endpoint(self, value: str) -> str:
        endpoint = value.strip()
        if "://" not in endpoint:
            endpoint = f"https://{endpoint}"
        if not endpoint.lower().endswith("/ews/exchange.asmx"):
            endpoint = endpoint.rstrip("/") + "/EWS/Exchange.asmx"
        return endpoint

    def _configure_ssl_verification(self) -> None:
        global _RAW_SESSION_PATCHED
        if _RAW_SESSION_PATCHED:
            return

        verify_ssl = self.settings.exchange_verify_ssl
        original = BaseProtocol.raw_session.__func__

        def raw_session_with_verify(cls, prefix, oauth2_client=None, oauth2_session_params=None, oauth2_token_endpoint=None):
            session = original(
                cls,
                prefix,
                oauth2_client=oauth2_client,
                oauth2_session_params=oauth2_session_params,
                oauth2_token_endpoint=oauth2_token_endpoint,
            )
            session.verify = verify_ssl
            return session

        BaseProtocol.raw_session = classmethod(raw_session_with_verify)
        _RAW_SESSION_PATCHED = True

        if not verify_ssl:
            warnings.filterwarnings("ignore", category=InsecureRequestWarning)
            logger.warning("SSL certificate verification is disabled for Exchange connections")

    def _configure_timezone_fallback(self) -> None:
        global _TIMEZONE_FALLBACK_PATCHED
        if _TIMEZONE_FALLBACK_PATCHED:
            return

        fallback_timezone = self.settings.exchange_timezone
        original = EWSTimeZone.from_ms_id.__func__

        def from_ms_id_with_fallback(cls, ms_id):
            try:
                return original(cls, ms_id)
            except UnknownTimeZone:
                if isinstance(ms_id, str) and _GUID_TIMEZONE_RE.match(ms_id):
                    logger.info(
                        "Mapping unknown Exchange timezone id %s to configured timezone %s",
                        ms_id,
                        fallback_timezone,
                    )
                    return cls(fallback_timezone)
                raise

        EWSTimeZone.from_ms_id = classmethod(from_ms_id_with_fallback)
        _TIMEZONE_FALLBACK_PATCHED = True

    def _resolve_folder(self, value: str | None) -> Folder:
        account = self.account
        if not value or value == "root":
            return account.root
        normalized = value.strip("/").lower()
        archive_root = account.root
        if normalized == "archive":
            try:
                archive_root = account.archive_root
            except Exception:  # noqa: BLE001
                archive_root = account.root
        builtin = {
            "inbox": account.inbox,
            "sent": account.sent,
            "sentitems": account.sent,
            "drafts": account.drafts,
            "deleted": account.trash,
            "trash": account.trash,
            "junk": account.junk,
            "archive": archive_root,
            "calendar": account.calendar,
            "contacts": account.contacts,
        }
        if normalized in builtin:
            return builtin[normalized]
        for folder in account.root.walk():
            if getattr(folder, "id", None) == value:
                return folder

        current = account.root
        for part in [segment for segment in value.strip("/").split("/") if segment]:
            next_folder = next(
                (child for child in current.children if child.name.lower() == part.lower()),
                None,
            )
            if next_folder is None:
                raise NotFoundError(value)
            current = next_folder
        return current

    def _fetch_item(self, item_id: str, folder: Folder | None = None) -> Any:
        try:
            return next(self.account.fetch(ids=[ItemId(id=item_id, changekey=None)], folder=folder))
        except StopIteration as exc:
            raise NotFoundError(item_id) from exc
        except Exception as exc:  # noqa: BLE001
            raise self._map_exception(exc, item_id=item_id) from exc

    def _mailbox(self, address: str) -> Mailbox:
        return Mailbox(email_address=address)

    def _email_address(self, mailbox: Any) -> EmailAddress:
        if mailbox is None:
            return EmailAddress(email="unknown@example.invalid", name=None)
        email = getattr(mailbox, "email_address", None) or getattr(mailbox, "email", None) or "unknown@example.invalid"
        name = getattr(mailbox, "name", None)
        return EmailAddress(email=email, name=name)

    def _recipients(self, values: Iterable[Any] | None) -> list[EmailAddress]:
        return [self._email_address(value) for value in values or []]

    def _to_email_summary(self, item: Any) -> EmailSummary:
        return EmailSummary(
            id=item.id,
            subject=item.subject or "",
            **{"from": self._email_address(getattr(item, "author", None) or getattr(item, "sender", None))},
            to=self._recipients(getattr(item, "to_recipients", None)),
            date=getattr(item, "datetime_received", None)
            or getattr(item, "datetime_sent", None)
            or getattr(item, "datetime_created", None)
            or datetime.now(UTC),
            is_read=bool(getattr(item, "is_read", False)),
            has_attachments=bool(getattr(item, "has_attachments", False)),
            preview=self._preview(item),
            importance=self._normalize_importance(getattr(item, "importance", None)),
            categories=list(getattr(item, "categories", None) or []),
        )

    def _attachment_metadata(self, attachment: Any) -> Attachment:
        return Attachment(
            id=getattr(getattr(attachment, "attachment_id", None), "id", None),
            name=getattr(attachment, "name", "attachment"),
            size=getattr(attachment, "size", None),
            content_type=getattr(attachment, "content_type", None),
        )

    def _to_email_full(self, item: Any) -> EmailFull:
        body_text, body_html = self._extract_message_body(item)
        return EmailFull(
            **self._to_email_summary(item).model_dump(by_alias=True),
            cc=self._recipients(getattr(item, "cc_recipients", None)),
            bcc=self._recipients(getattr(item, "bcc_recipients", None)),
            body_text=body_text,
            body_html=body_html,
            attachments=[self._attachment_metadata(a) for a in getattr(item, "attachments", None) or []],
            conversation_id=getattr(getattr(item, "conversation_id", None), "id", None),
            headers=self._headers_to_dict(getattr(item, "headers", None)),
            truncated=False,
        )

    def _extract_message_body(self, item: Any) -> tuple[str, str | None]:
        text = ""
        html = None
        if getattr(item, "text_body", None):
            text = str(item.text_body)
        elif getattr(item, "body", None):
            text = str(item.body)

        body = getattr(item, "body", None)
        if isinstance(body, HTMLBody):
            html = str(body)
        elif body is not None and "</" in str(body):
            html = str(body)
        return text, html

    def _preview(self, item: Any) -> str:
        text, _ = self._extract_message_body(item)
        return text[:200]

    def _headers_to_dict(self, headers: Any) -> dict[str, str]:
        result: dict[str, str] = {}
        for header in headers or []:
            name = getattr(header, "name", None)
            value = getattr(header, "value", None)
            if name and value is not None:
                result[str(name)] = str(value)
        return result

    def _normalize_importance(self, value: Any) -> str:
        normalized = str(value or "normal").lower()
        return normalized if normalized in {"low", "normal", "high"} else "normal"

    def _to_calendar_event(self, item: Any) -> CalendarEvent:
        attendees = [
            self._to_attendee(attendee)
            for attendee in (getattr(item, "required_attendees", None) or []) + (getattr(item, "optional_attendees", None) or [])
        ]
        organizer = self._email_address(getattr(item, "organizer", None))
        body_text, _ = self._extract_message_body(item)
        online_url = getattr(item, "meeting_workspace_url", None) or getattr(item, "net_show_url", None)
        recurrence = getattr(item, "recurrence", None)
        return CalendarEvent(
            id=item.id,
            subject=item.subject or "",
            start=item.start,
            end=item.end,
            location=getattr(item, "location", None),
            organizer=organizer,
            attendees=attendees,
            is_all_day=bool(getattr(item, "is_all_day", False)),
            is_recurring=bool(getattr(item, "is_recurring", False)),
            my_response=self._normalize_response(getattr(item, "my_response_type", None)),
            online_meeting_url=str(online_url) if online_url else None,
            body=body_text or None,
            reminder_minutes=getattr(item, "reminder_minutes_before_start", None),
            categories=list(getattr(item, "categories", None) or []),
            recurrence_pattern={"value": str(recurrence)} if recurrence else None,
            importance=self._normalize_importance(getattr(item, "importance", None)),
        )

    def _to_attendee(self, attendee: Any) -> Any:
        mailbox = getattr(attendee, "mailbox", None) or attendee
        response = getattr(attendee, "response_type", None)
        address = self._email_address(mailbox)
        from .models import Attendee as ApiAttendee

        return ApiAttendee(
            email=address.email,
            name=address.name,
            response_type=self._normalize_response(response),
        )

    def _normalize_response(self, value: Any) -> str:
        normalized = str(value or "unknown").lower()
        mapping = {
            "accept": "accept",
            "accepted": "accept",
            "organizer": "accept",
            "tentative": "tentative",
            "decline": "decline",
            "declined": "decline",
        }
        return mapping.get(normalized, "unknown")

    def _make_message(self, request: SendEmailRequest | DraftEmailRequest) -> Message:
        body: str | HTMLBody = HTMLBody(request.body) if request.body_type == "html" else request.body
        message = Message(
            account=self.account,
            folder=self.account.drafts,
            subject=request.subject,
            body=body,
            to_recipients=[self._mailbox(address) for address in request.to],
            cc_recipients=[self._mailbox(address) for address in request.cc],
            bcc_recipients=[self._mailbox(address) for address in request.bcc],
            reply_to=[self._mailbox(request.reply_to)] if getattr(request, "reply_to", None) else None,
            importance=request.importance.capitalize() if hasattr(request, "importance") else "Normal",
        )
        for path in request.attachments:
            with Path(path).open("rb") as handle:
                message.attach(FileAttachment(name=Path(path).name, content=handle.read()))
        return message

    def _to_folder_info(self, folder: Folder, depth: int) -> FolderInfo:
        children = []
        if depth > 0:
            children = [self._to_folder_info(child, depth - 1) for child in folder.children]
        return FolderInfo(
            id=getattr(folder, "id", None),
            name=folder.name,
            path=self._folder_path(folder),
            unread_count=getattr(folder, "unread_count", 0) or 0,
            total_count=getattr(folder, "total_count", 0) or 0,
            children=children,
        )

    def _folder_path(self, folder: Folder) -> str:
        parts = []
        current = folder
        while current is not None and getattr(current, "name", None):
            parts.append(current.name)
            current = getattr(current, "parent", None)
        return "/".join(reversed(parts))

    def _map_exception(self, exc: Exception, item_id: str | None = None) -> APIError:
        message = str(exc)
        if isinstance(exc, APIError):
            return exc
        if isinstance(exc, UnauthorizedError):
            return AuthFailedError()
        if isinstance(exc, RateLimitError):
            return ExchangeUnavailableError("exchange throttling or rate limit encountered")
        if isinstance(exc, (TransportError, TimeoutError)):
            if "timed out" in message.lower():
                return TimeoutAPIError(self.settings.exchange_timeout)
            return ExchangeUnavailableError(message)
        if isinstance(exc, (ErrorItemSavePropertyError, ErrorFolderSavePropertyError)):
            return ConflictError(message)
        if isinstance(exc, ResponseMessageError):
            lowered = message.lower()
            if "not found" in lowered and item_id:
                return NotFoundError(item_id)
            if "access is denied" in lowered or "permission" in lowered:
                return PermissionDeniedError()
            return APIError("exchange_error", message)
        return ExchangeUnavailableError(message)

    def ping(self) -> PingResult:
        started = datetime.now(UTC)
        account = self.account
        try:
            inbox = account.inbox
            count = inbox.total_count
            del count
        except Exception as exc:  # noqa: BLE001
            raise self._map_exception(exc) from exc
        latency_ms = round((datetime.now(UTC) - started).total_seconds() * 1000)
        parsed = urlparse(self.settings.exchange_server)
        version = getattr(getattr(account.protocol, "version", None), "api_version", None)
        return PingResult(status="ok", server=parsed.netloc or self.settings.exchange_server, version=version, latency_ms=latency_ms)

    def get_mailbox_info(self) -> MailboxInfo:
        account = self.account
        version = getattr(getattr(account.protocol, "version", None), "api_version", None)
        return MailboxInfo(
            email_address=account.primary_smtp_address,
            display_name=account.fullname or account.primary_smtp_address,
            timezone=str(account.default_timezone),
            exchange_version=version,
        )

    def list_emails(self, request: ListEmailsRequest) -> list[EmailSummary]:
        folder = self._resolve_folder(request.folder)
        qs = folder.all().order_by("-datetime_received")
        filters: dict[str, Any] = {}
        if request.from_address:
            filters["author__email_address"] = str(request.from_address)
        if request.subject:
            filters["subject__icontains"] = request.subject
        if request.since:
            filters["datetime_received__gte"] = datetime.combine(request.since, datetime.min.time(), tzinfo=self.account.default_timezone)
        if request.before:
            filters["datetime_received__lt"] = datetime.combine(request.before + timedelta(days=1), datetime.min.time(), tzinfo=self.account.default_timezone)
        if request.unread_only:
            filters["is_read"] = False
        if request.has_attachments is not None:
            filters["has_attachments"] = request.has_attachments
        if filters:
            qs = qs.filter(**filters)
        try:
            items = list(qs[request.offset : request.offset + request.limit])
        except Exception as exc:  # noqa: BLE001
            raise self._map_exception(exc) from exc
        return [self._to_email_summary(item) for item in items]

    def get_email(self, request: GetEmailRequest) -> EmailFull:
        item = self._fetch_item(request.id)
        return self._to_email_full(item)

    def search_emails(self, request: SearchEmailsRequest) -> list[EmailSummary]:
        folder = self._resolve_folder(request.folder) if request.folder else self.account.inbox
        try:
            qs = folder.filter(subject__icontains=request.query).order_by("-datetime_received")
            items = list(qs[: request.limit])
        except Exception:
            items = []
        if not items:
            try:
                qs = folder.filter(text_body__icontains=request.query).order_by("-datetime_received")
                items = list(qs[: request.limit])
            except Exception as exc:  # noqa: BLE001
                raise self._map_exception(exc) from exc
        return [self._to_email_summary(item) for item in items]

    def send_email(self, request: SendEmailRequest) -> SendResult:
        message = self._make_message(request)
        try:
            message.send_and_save()
            return SendResult(id=message.id or "", status="sent")
        except Exception as exc:  # noqa: BLE001
            raise self._map_exception(exc) from exc

    def reply_email(self, request: ReplyEmailRequest) -> SendResult:
        item = self._fetch_item(request.id)
        try:
            if request.reply_all:
                item.reply_all(subject=f"Re: {item.subject or ''}", body=request.body)
            else:
                item.reply(subject=f"Re: {item.subject or ''}", body=request.body)
            return SendResult(id=request.id, status="sent")
        except Exception as exc:  # noqa: BLE001
            raise self._map_exception(exc, item_id=request.id) from exc

    def forward_email(self, request: ForwardEmailRequest) -> SendResult:
        item = self._fetch_item(request.id)
        try:
            item.forward(
                subject=f"Fwd: {item.subject or ''}",
                body=request.comment or "",
                to_recipients=[self._mailbox(address) for address in request.to],
            )
            return SendResult(id=request.id, status="sent")
        except Exception as exc:  # noqa: BLE001
            raise self._map_exception(exc, item_id=request.id) from exc

    def move_email(self, request: FolderActionRequest) -> ActionResult:
        item = self._fetch_item(request.id)
        destination = self._resolve_folder(request.folder)
        try:
            result = item.move(to_folder=destination)
            return ActionResult(id=getattr(result, "id", request.id), status="moved", new_folder=request.folder)
        except Exception as exc:  # noqa: BLE001
            raise self._map_exception(exc, item_id=request.id) from exc

    def copy_email(self, request: FolderActionRequest) -> ActionResult:
        item = self._fetch_item(request.id)
        destination = self._resolve_folder(request.folder)
        try:
            result = item.copy(to_folder=destination)
            return ActionResult(
                id=request.id,
                status="copied",
                new_folder=request.folder,
                new_id=getattr(result, "id", None),
            )
        except Exception as exc:  # noqa: BLE001
            raise self._map_exception(exc, item_id=request.id) from exc

    def delete_email(self, request: DeleteEmailRequest) -> ActionResult:
        item = self._fetch_item(request.id)
        delete_type = "HardDelete" if request.hard_delete else "MoveToDeletedItems"
        try:
            item.delete(delete_type=delete_type)
            return ActionResult(id=request.id, status="deleted")
        except Exception as exc:  # noqa: BLE001
            raise self._map_exception(exc, item_id=request.id) from exc

    def mark_email(self, request: MarkEmailRequest) -> ActionResult:
        item = self._fetch_item(request.id)
        updated_fields: list[str] = []
        if request.read is not None:
            item.is_read = request.read
            updated_fields.append("read")
        if request.importance is not None:
            item.importance = request.importance.capitalize()
            updated_fields.append("importance")
        if request.flag is not None:
            item.categories = [] if request.flag == "none" else [request.flag]
            updated_fields.append("flag")
        try:
            item.save(update_fields=updated_fields or None)
            return ActionResult(id=request.id, status="updated", updated_fields=updated_fields)
        except Exception as exc:  # noqa: BLE001
            raise self._map_exception(exc, item_id=request.id) from exc

    def list_folders(self, request: ListFoldersRequest) -> list[FolderInfo]:
        folder = self._resolve_folder(request.parent)
        return [self._to_folder_info(child, request.depth - 1) for child in folder.children] if request.depth != 0 else [self._to_folder_info(folder, 0)]

    def create_folder(self, request: CreateFolderRequest) -> ActionResult:
        parent = self._resolve_folder(request.parent)
        folder = Folder(parent=parent, name=request.name)
        try:
            folder.save()
            return ActionResult(id=getattr(folder, "id", ""), status="created", path=self._folder_path(folder))
        except Exception as exc:  # noqa: BLE001
            raise self._map_exception(exc) from exc

    def create_draft(self, request: DraftEmailRequest) -> ActionResult:
        message = self._make_message(request)
        try:
            message.save()
            return ActionResult(id=message.id or "", status="draft")
        except Exception as exc:  # noqa: BLE001
            raise self._map_exception(exc) from exc

    def send_draft(self, request: SendDraftRequest) -> ActionResult:
        item = self._fetch_item(request.id, folder=self.account.drafts)
        try:
            item.send_and_save()
            return ActionResult(id=request.id, status="sent")
        except Exception as exc:  # noqa: BLE001
            raise self._map_exception(exc, item_id=request.id) from exc

    def get_attachment(self, request: GetAttachmentRequest) -> AttachmentResult:
        item = self._fetch_item(request.email_id)
        target_dir = Path(request.save_path) if request.save_path else Path(tempfile.gettempdir())
        target_dir.mkdir(parents=True, exist_ok=True)
        for attachment in getattr(item, "attachments", None) or []:
            attachment_id = getattr(getattr(attachment, "attachment_id", None), "id", None)
            if attachment_id == request.attachment_id:
                filename = getattr(attachment, "name", "attachment.bin")
                path = self._unique_path(target_dir / filename)
                content = getattr(attachment, "content", None)
                if content is None:
                    _ = attachment.content
                    content = attachment.content
                path.write_bytes(content)
                return AttachmentResult(
                    filename=filename,
                    size=len(content),
                    saved_path=str(path),
                    content_type=getattr(attachment, "content_type", None),
                )
        raise NotFoundError(request.attachment_id)

    def _unique_path(self, path: Path) -> Path:
        if not path.exists():
            return path
        stem = path.stem
        suffix = path.suffix
        index = 1
        while True:
            candidate = path.with_name(f"{stem}-{index}{suffix}")
            if not candidate.exists():
                return candidate
            index += 1

    def list_events(self, request: ListEventsRequest) -> list[CalendarEvent]:
        folder = self.account.calendar if not request.calendar_id else self._resolve_folder(request.calendar_id)
        qs = folder.view(start=request.start, end=request.end)
        try:
            items = list(qs)
        except Exception as exc:  # noqa: BLE001
            raise self._map_exception(exc) from exc
        return [self._to_calendar_event(item) for item in items]

    def get_event(self, request: GetEventRequest) -> CalendarEvent:
        item = self._fetch_item(request.id, folder=self.account.calendar)
        return self._to_calendar_event(item)

    def create_event(self, request: CreateEventRequest) -> CreateEventResult:
        folder = self.account.calendar if not request.calendar_id else self._resolve_folder(request.calendar_id)
        item = CalendarItem(
            account=self.account,
            folder=folder,
            subject=request.subject,
            start=request.start,
            end=request.end,
            location=request.location,
            body=request.body,
            required_attendees=[Attendee(mailbox=self._mailbox(address)) for address in request.attendees],
            is_all_day=request.is_all_day,
            categories=request.categories,
            importance=request.importance.capitalize(),
            reminder_minutes_before_start=request.reminder_minutes,
        )
        try:
            item.save(send_meeting_invitations="SendToAllAndSaveCopy" if request.attendees else "SendToNone")
            return CreateEventResult(
                id=item.id or "",
                status="created",
                subject=request.subject,
                start=request.start,
                end=request.end,
                invite_sent=bool(request.attendees),
            )
        except Exception as exc:  # noqa: BLE001
            raise self._map_exception(exc) from exc

    def update_event(self, request: UpdateEventRequest) -> ActionResult:
        item = self._fetch_item(request.id, folder=self.account.calendar)
        updated_fields: list[str] = []
        for field in ["subject", "start", "end", "location", "body", "reminder_minutes"]:
            value = getattr(request, field)
            if value is not None:
                target = "reminder_minutes_before_start" if field == "reminder_minutes" else field
                setattr(item, target, value)
                updated_fields.append(field)
        if request.add_attendees:
            current = list(getattr(item, "required_attendees", None) or [])
            current.extend(Attendee(mailbox=self._mailbox(address)) for address in request.add_attendees)
            item.required_attendees = current
            updated_fields.append("add_attendees")
        if request.remove_attendees:
            remove_set = {address.lower() for address in request.remove_attendees}
            item.required_attendees = [
                attendee
                for attendee in getattr(item, "required_attendees", None) or []
                if getattr(getattr(attendee, "mailbox", None), "email_address", "").lower() not in remove_set
            ]
            updated_fields.append("remove_attendees")
        try:
            invitations = {
                "none": "SendToNone",
                "all": "SendToAllAndSaveCopy",
                "modified": "SendOnlyToChanged",
            }[request.send_updates]
            item.save(update_fields=updated_fields or None, send_meeting_invitations=invitations)
            return ActionResult(id=request.id, status="updated", updated_fields=updated_fields)
        except Exception as exc:  # noqa: BLE001
            raise self._map_exception(exc, item_id=request.id) from exc

    def delete_event(self, request: DeleteEventRequest) -> ActionResult:
        item = self._fetch_item(request.id, folder=self.account.calendar)
        try:
            item.delete(
                send_meeting_cancellations="SendToAllAndSaveCopy" if request.notify_attendees else "SendToNone"
            )
            return ActionResult(id=request.id, status="deleted")
        except Exception as exc:  # noqa: BLE001
            raise self._map_exception(exc, item_id=request.id) from exc

    def respond_to_invite(self, request: RespondToInviteRequest) -> ActionResult:
        item = self._fetch_item(request.id, folder=self.account.calendar)
        try:
            if request.response == "accept":
                item.accept(body=request.message)
            elif request.response == "tentative":
                item.tentatively_accept(body=request.message)
            else:
                item.decline(body=request.message)
            return ActionResult(id=request.id, status=request.response)
        except Exception as exc:  # noqa: BLE001
            raise self._map_exception(exc, item_id=request.id) from exc

    def find_free_slots(self, request: FindFreeSlotsRequest) -> list[FreeSlot]:
        tz = self.account.default_timezone
        service = GetUserAvailability(protocol=self.account.protocol)
        mailbox_data = [
            MailboxData(email=self._mailbox(address), attendee_type="Required", exclude_conflicts=False)
            for address in request.attendees
        ]
        try:
            views = service.call(
                tzinfo=tz,
                mailbox_data=mailbox_data,
                timezone=tz,
                free_busy_view_options=FreeBusyViewOptions(
                    time_window=TimeWindow(start=request.start, end=request.end),
                    merged_free_busy_interval=request.duration,
                    requested_view="DetailedMerged",
                ),
            )
        except Exception as exc:  # noqa: BLE001
            raise self._map_exception(exc) from exc

        if not views:
            return []
        merged = getattr(views[0], "merged", "") or ""
        slots: list[FreeSlot] = []
        cursor = request.start
        interval = timedelta(minutes=request.duration)
        for symbol in merged:
            slot_end = cursor + interval
            if symbol == "0":
                slots.append(FreeSlot(start=cursor, end=slot_end, all_available=True, busy_attendees=[]))
            cursor = slot_end
            if cursor >= request.end:
                break
        return slots

    def get_my_availability(self, request: ListEventsRequest) -> AvailabilityResult:
        events = self.list_events(request)
        busy_slots = [{"start": event.start, "end": event.end, "subject": event.subject} for event in events]
        return AvailabilityResult(free_slots=[], busy_slots=busy_slots)

    def list_calendars(self) -> list[CalendarInfo]:
        return [
            CalendarInfo(
                id=self.account.calendar.id,
                name=self.account.calendar.name,
                is_default=True,
                owner_email=self.account.primary_smtp_address,
            )
        ]

    def _contact_summary_from_contact(self, contact: Contact, source: str) -> ContactSummary:
        emails = [entry.email for entry in getattr(contact, "email_addresses", None) or [] if getattr(entry, "email", None)]
        phones = [entry.phone_number for entry in getattr(contact, "phone_numbers", None) or [] if getattr(entry, "phone_number", None)]
        return ContactSummary(
            id=contact.id,
            display_name=contact.display_name or contact.file_as or "",
            email_addresses=emails,
            phone_numbers=phones,
            company=getattr(contact, "company_name", None),
            job_title=getattr(contact, "job_title", None),
            department=getattr(contact, "department", None),
            source=source,
        )

    def search_contacts(self, request: SearchContactsRequest) -> list[ContactSummary]:
        results: list[ContactSummary] = []
        if request.source in {"personal", "all"}:
            try:
                qs = self.account.contacts.filter(display_name__icontains=request.query)[: request.limit]
                results.extend(self._contact_summary_from_contact(contact, "personal") for contact in qs)
            except Exception as exc:  # noqa: BLE001
                raise self._map_exception(exc) from exc
        if request.source in {"gal", "all"} and len(results) < request.limit:
            try:
                resolved = ResolveNames(protocol=self.account.protocol).call(
                    unresolved_entries=[request.query],
                    return_full_contact_data=True,
                    search_scope="ActiveDirectory",
                    contact_data_shape="AllProperties",
                )
                for mailbox, contact in resolved:
                    if contact is not None and getattr(contact, "id", None):
                        results.append(self._contact_summary_from_contact(contact, "gal"))
                    elif mailbox is not None:
                        results.append(
                            ContactSummary(
                                id=getattr(mailbox, "email_address", None) or request.query,
                                display_name=getattr(mailbox, "name", None) or getattr(mailbox, "email_address", None) or request.query,
                                email_addresses=[getattr(mailbox, "email_address", None)] if getattr(mailbox, "email_address", None) else [],
                                phone_numbers=[],
                                source="gal",
                            )
                        )
                    if len(results) >= request.limit:
                        break
            except Exception as exc:  # noqa: BLE001
                raise self._map_exception(exc) from exc
        return results[: request.limit]

    def get_contact(self, request: GetContactRequest) -> ContactFull:
        item = self._fetch_item(request.id, folder=self.account.contacts)
        return ContactFull(
            id=item.id,
            display_name=item.display_name or item.file_as or "",
            first_name=getattr(item, "given_name", None),
            last_name=getattr(item, "surname", None),
            email_addresses=[
                {"type": entry.label, "address": entry.email}
                for entry in getattr(item, "email_addresses", None) or []
                if getattr(entry, "email", None)
            ],
            phone_numbers=[
                {"type": entry.label, "number": entry.phone_number}
                for entry in getattr(item, "phone_numbers", None) or []
                if getattr(entry, "phone_number", None)
            ],
            addresses=[
                {
                    "type": entry.label,
                    "street": entry.street,
                    "city": entry.city,
                    "state": entry.state,
                    "postal_code": entry.zipcode,
                    "country": entry.country,
                }
                for entry in getattr(item, "physical_addresses", None) or []
            ],
            company=getattr(item, "company_name", None),
            job_title=getattr(item, "job_title", None),
            department=getattr(item, "department", None),
            manager=getattr(item, "manager", None),
            notes=getattr(item, "notes", None),
            birthday=getattr(item, "birthday", None),
            source="personal",
        )

    def create_contact(self, request: CreateContactRequest) -> ActionResult:
        contact = Contact(
            account=self.account,
            folder=self.account.contacts,
            display_name=request.display_name,
            given_name=request.first_name,
            surname=request.last_name,
            company_name=request.company,
            job_title=request.job_title,
            notes=request.notes,
            email_addresses=[IndexedEmailAddress(label="EmailAddress1", email=str(request.email))] if request.email else [],
            phone_numbers=[IndexedPhoneNumber(label="PrimaryPhone", phone_number=request.phone)] if request.phone else [],
        )
        try:
            contact.save()
            return ActionResult(id=contact.id or "", status="created")
        except Exception as exc:  # noqa: BLE001
            raise self._map_exception(exc) from exc

    def update_contact(self, request: UpdateContactRequest) -> ActionResult:
        contact = self._fetch_item(request.id, folder=self.account.contacts)
        updated_fields: list[str] = []
        field_map = {
            "display_name": "display_name",
            "first_name": "given_name",
            "last_name": "surname",
            "company": "company_name",
            "job_title": "job_title",
            "notes": "notes",
        }
        for request_field, item_field in field_map.items():
            value = getattr(request, request_field)
            if value is not None:
                setattr(contact, item_field, value)
                updated_fields.append(request_field)
        if request.email is not None:
            contact.email_addresses = [IndexedEmailAddress(label="EmailAddress1", email=str(request.email))]
            updated_fields.append("email")
        if request.phone is not None:
            contact.phone_numbers = [IndexedPhoneNumber(label="PrimaryPhone", phone_number=request.phone)]
            updated_fields.append("phone")
        try:
            contact.save(update_fields=updated_fields or None)
            return ActionResult(id=request.id, status="updated", updated_fields=updated_fields)
        except Exception as exc:  # noqa: BLE001
            raise self._map_exception(exc, item_id=request.id) from exc

    def delete_contact(self, request: DeleteContactRequest) -> ActionResult:
        contact = self._fetch_item(request.id, folder=self.account.contacts)
        try:
            contact.delete(delete_type="MoveToDeletedItems")
            return ActionResult(id=request.id, status="deleted")
        except Exception as exc:  # noqa: BLE001
            raise self._map_exception(exc, item_id=request.id) from exc
