from __future__ import annotations

from ..config import Settings
from ..models import (
    ActionResult,
    AttachmentResult,
    AvailabilityResult,
    CalendarInfo,
    CalendarEvent,
    ContactFull,
    ContactSummary,
    CategorizeEmailRequest,
    CategoryUsage,
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
    FolderActionRequest,
    FolderInfo,
    FindFreeSlotsRequest,
    ForwardEmailRequest,
    FreeSlot,
    GetAttachmentRequest,
    GetContactRequest,
    GetEmailRequest,
    GetEventRequest,
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
    UpdateContactRequest,
    UpdateEventRequest,
)
from .protocol import ExchangeBackend
from .unconfigured import UnconfiguredExchangeBackend


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

    def categorize_email(self, request: CategorizeEmailRequest) -> ActionResult:
        return self.backend.categorize_email(request)

    def list_categories(self, request: ListCategoriesRequest) -> list[CategoryUsage]:
        return self.backend.list_categories(request)

    def list_folders(self, request: ListFoldersRequest) -> list[FolderInfo]:
        return self.backend.list_folders(request)

    def create_folder(self, request: CreateFolderRequest) -> ActionResult:
        return self.backend.create_folder(request)

    def create_draft(self, request: DraftEmailRequest) -> ActionResult:
        return self.backend.create_draft(request)

    def send_draft(self, request: SendDraftRequest) -> SendResult:
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
