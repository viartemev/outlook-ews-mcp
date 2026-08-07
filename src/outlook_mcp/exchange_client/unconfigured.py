from __future__ import annotations

from typing import NoReturn

from ..config import Settings
from ..errors import ExchangeUnavailableError
from ..models import (
    ActionResult,
    AttachmentResult,
    AvailabilityResult,
    BulkCategorizeEmailsRequest,
    BulkDeleteEmailsRequest,
    BulkMarkEmailsRequest,
    BulkMoveEmailsRequest,
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
    CreateRuleRequest,
    DeleteContactRequest,
    DeleteEmailRequest,
    DeleteEventRequest,
    DeleteFolderRequest,
    DeleteRuleRequest,
    DraftEmailRequest,
    EmailFull,
    EmailMimeResult,
    EmailSummary,
    FolderActionRequest,
    FolderInfo,
    FindFreeSlotsRequest,
    ForwardEmailRequest,
    FreeSlot,
    GetAttachmentRequest,
    GetContactRequest,
    GetEmailMimeRequest,
    GetEmailRequest,
    GetEventRequest,
    GetThreadRequest,
    ListCategoriesRequest,
    ListEmailsRequest,
    ListEventsRequest,
    ListFoldersRequest,
    ListRoomsRequest,
    MailboxInfo,
    MailRule,
    MarkEmailRequest,
    OofSettingsModel,
    PingResult,
    RenameFolderRequest,
    ReplyEmailRequest,
    RespondToInviteRequest,
    RoomInfo,
    RoomListInfo,
    SearchContactsRequest,
    SearchEmailsRequest,
    SendDraftRequest,
    SendEmailRequest,
    SendResult,
    SetOofSettingsRequest,
    Thread,
    UpdateContactRequest,
    UpdateDraftRequest,
    UpdateEventRequest,
    UpdateRuleRequest,
)


class UnconfiguredExchangeBackend:
    def __init__(self, settings: Settings) -> None:
        self.settings = settings

    def _raise(self) -> NoReturn:
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

    def get_thread(self, request: GetThreadRequest) -> Thread:
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

    def categorize_email(self, request: CategorizeEmailRequest) -> ActionResult:
        self._raise()

    def bulk_move_emails(self, request: BulkMoveEmailsRequest) -> list[ActionResult]:
        self._raise()

    def bulk_delete_emails(self, request: BulkDeleteEmailsRequest) -> list[ActionResult]:
        self._raise()

    def bulk_mark_emails(self, request: BulkMarkEmailsRequest) -> list[ActionResult]:
        self._raise()

    def bulk_categorize_emails(self, request: BulkCategorizeEmailsRequest) -> list[ActionResult]:
        self._raise()

    def list_categories(self, request: ListCategoriesRequest) -> list[CategoryUsage]:
        self._raise()

    def list_folders(self, request: ListFoldersRequest) -> list[FolderInfo]:
        self._raise()

    def create_folder(self, request: CreateFolderRequest) -> ActionResult:
        self._raise()

    def rename_folder(self, request: RenameFolderRequest) -> ActionResult:
        self._raise()

    def delete_folder(self, request: DeleteFolderRequest) -> ActionResult:
        self._raise()

    def create_draft(self, request: DraftEmailRequest) -> ActionResult:
        self._raise()

    def update_draft(self, request: UpdateDraftRequest) -> ActionResult:
        self._raise()

    def send_draft(self, request: SendDraftRequest) -> SendResult:
        self._raise()

    def get_attachment(self, request: GetAttachmentRequest) -> AttachmentResult:
        self._raise()

    def get_email_mime(self, request: GetEmailMimeRequest) -> EmailMimeResult:
        self._raise()

    def list_rules(self) -> list[MailRule]:
        self._raise()

    def create_rule(self, request: CreateRuleRequest) -> ActionResult:
        self._raise()

    def update_rule(self, request: UpdateRuleRequest) -> ActionResult:
        self._raise()

    def delete_rule(self, request: DeleteRuleRequest) -> ActionResult:
        self._raise()

    def get_oof_settings(self) -> OofSettingsModel:
        self._raise()

    def set_oof_settings(self, request: SetOofSettingsRequest) -> ActionResult:
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

    def list_room_lists(self) -> list[RoomListInfo]:
        self._raise()

    def list_rooms(self, request: ListRoomsRequest) -> list[RoomInfo]:
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
