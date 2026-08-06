from __future__ import annotations

from datetime import UTC, datetime, timedelta
from pathlib import Path

import pytest

from outlook_mcp.config import Settings
from outlook_mcp.exchange_client import ExchangeClient
from outlook_mcp.models import (
    ActionResult,
    AttachmentResult,
    AvailabilityResult,
    CalendarInfo,
    CalendarEvent,
    ContactFull,
    ContactSummary,
    CreateContactRequest,
    CategorizeEmailRequest,
    CategoryUsage,
    CreateEventRequest,
    CreateEventResult,
    DeleteContactRequest,
    DeleteEmailRequest,
    DeleteEventRequest,
    DraftEmailRequest,
    EmailAddress,
    EmailFull,
    EmailSummary,
    FolderActionRequest,
    FolderInfo,
    FindFreeSlotsRequest,
    ForwardEmailRequest,
    FreeSlot,
    GetAttachmentRequest,
    GetContactRequest,
    GetEmailRequest,
    GetEventRequest,
    GetThreadRequest,
    ListCategoriesRequest,
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
    Thread,
    UpdateContactRequest,
    UpdateEventRequest,
)


class FakeExchangeBackend:
    def ping(self) -> PingResult:
        return PingResult(
            status="ok",
            server="https://mail.example.com/EWS/Exchange.asmx",
            version="2019",
            latency_ms=42,
        )

    def get_mailbox_info(self) -> MailboxInfo:
        return MailboxInfo(
            email_address="user@example.com",
            display_name="Test User",
            timezone="Europe/Moscow",
            mailbox_size_mb=256.5,
            quota_mb=1024,
            exchange_version="2019",
        )

    def list_emails(self, request: ListEmailsRequest) -> list[EmailSummary]:
        return [
            EmailSummary(
                id="email-1",
                subject=request.subject or "Hello",
                **{"from": {"email": "sender@example.com", "name": "Sender"}},
                to=[EmailAddress(email="user@example.com", name="User")],
                date=datetime(2026, 4, 7, 10, 0, tzinfo=UTC),
                is_read=not request.unread_only,
                has_attachments=bool(request.has_attachments),
                preview="preview",
            )
        ]

    def get_email(self, request: GetEmailRequest) -> EmailFull:
        return EmailFull(
            id=request.id,
            subject="Hello",
            **{"from": {"email": "sender@example.com", "name": "Sender"}},
            to=[EmailAddress(email="user@example.com", name="User")],
            date=datetime(2026, 4, 7, 10, 0, tzinfo=UTC),
            is_read=True,
            has_attachments=False,
            preview="preview",
            body_text="Body",
            conversation_id="conv-1",
            headers={"X-Test": "1"},
        )

    def get_thread(self, request: GetThreadRequest) -> Thread:
        messages = [
            EmailFull(
                id=f"email-{index}",
                subject="Hello" if index == 1 else f"Re: Hello ({index})",
                **{"from": {"email": "sender@example.com", "name": "Sender"}},
                to=[EmailAddress(email="user@example.com", name="User")],
                date=datetime(2026, 4, 7, 9 + index, 0, tzinfo=UTC),
                is_read=True,
                body_text=f"Body {index}",
                conversation_id="conv-1",
            )
            for index in (1, 2)
        ]
        return Thread(
            conversation_id=request.conversation_id or "conv-1",
            subject="Hello",
            message_count=len(messages),
            messages=messages,
        )

    def search_emails(self, request: SearchEmailsRequest) -> list[EmailSummary]:
        return self.list_emails(ListEmailsRequest(subject=request.query, limit=request.limit))

    def send_email(self, request: SendEmailRequest) -> SendResult:
        return SendResult(id="sent-1", status="sent")

    def reply_email(self, request: ReplyEmailRequest) -> SendResult:
        return SendResult(id=request.id, status="sent")

    def forward_email(self, request: ForwardEmailRequest) -> SendResult:
        return SendResult(id=request.id, status="sent")

    def move_email(self, request: FolderActionRequest) -> ActionResult:
        return ActionResult(id=request.id, status="moved", new_folder=request.folder)

    def copy_email(self, request: FolderActionRequest) -> ActionResult:
        return ActionResult(
            id=request.id,
            status="copied",
            new_folder=request.folder,
            new_id="email-copy-1",
        )

    def delete_email(self, request: DeleteEmailRequest) -> ActionResult:
        return ActionResult(id=request.id, status="deleted")

    def mark_email(self, request: MarkEmailRequest) -> ActionResult:
        updated_fields = [
            field for field in ["read", "flag", "importance"] if getattr(request, field) is not None
        ]
        return ActionResult(id=request.id, status="updated", updated_fields=updated_fields)

    def categorize_email(self, request: CategorizeEmailRequest) -> ActionResult:
        existing = ["Existing"]
        if request.mode == "set":
            categories = list(request.categories)
        elif request.mode == "add":
            categories = existing + [name for name in request.categories if name not in existing]
        else:
            categories = [name for name in existing if name not in request.categories]
        return ActionResult(
            id=request.id,
            status="updated",
            updated_fields=["categories"],
            categories=categories,
        )

    def list_categories(self, request: ListCategoriesRequest) -> list[CategoryUsage]:
        return [
            CategoryUsage(name="Important", count=7),
            CategoryUsage(name="Later", count=2),
        ]

    def list_folders(self, request: ListFoldersRequest) -> list[FolderInfo]:
        return [
            FolderInfo(
                id="folder-inbox",
                name="Inbox",
                path="inbox",
                unread_count=3,
                total_count=10,
                children=[],
            )
        ]

    def create_folder(self, request) -> ActionResult:
        return ActionResult(
            id="folder-new", status="created", path=f"{request.parent}/{request.name}"
        )

    def create_draft(self, request: DraftEmailRequest) -> ActionResult:
        return ActionResult(id="draft-1", status="draft")

    def send_draft(self, request: SendDraftRequest) -> SendResult:
        return SendResult(id=request.id, status="sent")

    def get_attachment(self, request: GetAttachmentRequest) -> AttachmentResult:
        target = request.save_path or Path("/tmp/test.txt")
        return AttachmentResult(
            filename="test.txt",
            size=5,
            saved_path=str(target),
            content_type="text/plain",
        )

    def list_events(self, request: ListEventsRequest) -> list[CalendarEvent]:
        return [
            CalendarEvent(
                id="event-1",
                subject="Planning",
                start=request.start,
                end=request.end,
                location="Room 1",
                organizer=EmailAddress(email="organizer@example.com", name="Organizer"),
            )
        ]

    def get_event(self, request: GetEventRequest) -> CalendarEvent:
        return CalendarEvent(
            id=request.id,
            subject="Planning",
            start=datetime(2026, 4, 8, 9, 0, tzinfo=UTC),
            end=datetime(2026, 4, 8, 10, 0, tzinfo=UTC),
            location="Room 1",
            organizer=EmailAddress(email="organizer@example.com", name="Organizer"),
        )

    def create_event(self, request: CreateEventRequest) -> CreateEventResult:
        return CreateEventResult(
            id="event-new",
            status="created",
            subject=request.subject,
            start=request.start,
            end=request.end,
            invite_sent=bool(request.attendees),
        )

    def update_event(self, request: UpdateEventRequest) -> ActionResult:
        updated_fields = [
            field
            for field in ["subject", "start", "end", "location", "body"]
            if getattr(request, field) is not None
        ]
        return ActionResult(id=request.id, status="updated", updated_fields=updated_fields)

    def delete_event(self, request: DeleteEventRequest) -> ActionResult:
        return ActionResult(id=request.id, status="deleted")

    def respond_to_invite(self, request: RespondToInviteRequest) -> ActionResult:
        return ActionResult(id=request.id, status=request.response)

    def find_free_slots(self, request: FindFreeSlotsRequest) -> list[FreeSlot]:
        return [
            FreeSlot(
                start=request.start,
                end=request.start + timedelta(hours=1),
                all_available=True,
            )
        ]

    def get_my_availability(self, request: ListEventsRequest) -> AvailabilityResult:
        return AvailabilityResult(
            free_slots=[
                FreeSlot(start=request.start, end=request.end, all_available=True),
            ],
            busy_slots=[],
        )

    def list_calendars(self) -> list[CalendarInfo]:
        return [
            CalendarInfo(
                id="cal-1", name="Calendar", is_default=True, owner_email="user@example.com"
            )
        ]

    def search_contacts(self, request: SearchContactsRequest) -> list[ContactSummary]:
        return [
            ContactSummary(
                id="contact-1",
                display_name="Ivan Ivanov",
                email_addresses=["ivan@example.com"],
                phone_numbers=["+79990000000"],
                company="Example",
                job_title="Manager",
                department="Sales",
                source="gal",
            )
        ]

    def get_contact(self, request: GetContactRequest) -> ContactFull:
        return ContactFull(
            id=request.id,
            display_name="Ivan Ivanov",
            first_name="Ivan",
            last_name="Ivanov",
            email_addresses=[{"type": "work", "address": "ivan@example.com"}],
            phone_numbers=[{"type": "mobile", "number": "+79990000000"}],
            source="gal",
        )

    def create_contact(self, request: CreateContactRequest) -> ActionResult:
        return ActionResult(id="contact-new", status="created")

    def update_contact(self, request: UpdateContactRequest) -> ActionResult:
        return ActionResult(id=request.id, status="updated", updated_fields=["display_name"])

    def delete_contact(self, request: DeleteContactRequest) -> ActionResult:
        return ActionResult(id=request.id, status="deleted")


@pytest.fixture
def anyio_backend() -> str:
    return "asyncio"


@pytest.fixture
def settings() -> Settings:
    # _env_file=None keeps a developer's local .env out of the test run -- otherwise
    # their real EXCHANGE_* values could leak into test behavior and failure output.
    return Settings(
        _env_file=None,
        EXCHANGE_SERVER="https://mail.example.com/EWS/Exchange.asmx",
        EXCHANGE_USERNAME="DOMAIN\\user",
        EXCHANGE_PASSWORD="secret",
    )


@pytest.fixture
def client(settings: Settings) -> ExchangeClient:
    return ExchangeClient(settings=settings, backend=FakeExchangeBackend())
