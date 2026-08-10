from __future__ import annotations

import time
from collections.abc import Callable
from typing import TypeVar

from ..config import Settings
from ..errors import APIError
from ..models import (
    ActionResult,
    AddAttachmentRequest,
    AttachmentResult,
    BulkCategorizeEmailsRequest,
    BulkDeleteEventsRequest,
    BulkMarkEmailsRequest,
    BulkRespondToInvitesRequest,
    DeleteAttachmentRequest,
    BulkDeleteEmailsRequest,
    BulkMoveEmailsRequest,
    BulkResult,
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
    DeleteFolderRequest,
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
    DelegateInfo,
    CreateInboxRuleRequest,
    DeleteInboxRuleRequest,
    InboxRule,
    UpdateInboxRuleRequest,
    MarkEmailRequest,
    OutOfOfficeSettings,
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
    Thread,
    UpdateContactRequest,
    UpdateDraftRequest,
    UpdateEventRequest,
)
from .protocol import ExchangeBackend
from .unconfigured import UnconfiguredExchangeBackend

T = TypeVar("T")


class ExchangeClient:
    #: Starting backoff between retried reads, in seconds; doubles each attempt,
    #: capped by whatever remains of the EXCHANGE_MAX_RETRY_WAIT_SECONDS deadline.
    _RETRY_BASE_SECONDS = 2.0

    def __init__(self, settings: Settings, backend: ExchangeBackend | None = None) -> None:
        self.settings = settings
        self.backend = backend or UnconfiguredExchangeBackend(settings)

    def _retry_read(self, call: Callable[[], T]) -> T:
        """Retry a read-only backend call while Exchange reports itself busy.

        The account's retry_policy is FailFast (exchange_client/base.py), so every EWS
        service call raises on its first transient error instead of exchangelib
        retrying it internally with no cumulative time limit. This loop is what
        decides whether to try again, bounded by a wall-clock deadline built from
        EXCHANGE_MAX_RETRY_WAIT_SECONDS -- not by exchangelib's own max_wait, which
        only capped a single backoff step, not total retry time.

        Only read-only calls go through this. Writes never do: retrying a
        send/create/delete after an ambiguous failure risks repeating a side effect
        that already happened on the server before the error came back.
        """
        deadline = time.monotonic() + self.settings.exchange_max_retry_wait
        wait = self._RETRY_BASE_SECONDS
        while True:
            try:
                return call()
            except APIError as exc:
                if not exc.retryable:
                    raise
                remaining = deadline - time.monotonic()
                if remaining <= 0:
                    raise
                time.sleep(min(wait, remaining))
                wait *= 2

    def ping(self) -> PingResult:
        return self._retry_read(self.backend.ping)

    def get_mailbox_info(self) -> MailboxInfo:
        return self._retry_read(self.backend.get_mailbox_info)

    def list_delegates(self) -> list[DelegateInfo]:
        return self._retry_read(self.backend.list_delegates)

    def list_inbox_rules(self) -> list[InboxRule]:
        return self._retry_read(self.backend.list_inbox_rules)

    def create_inbox_rule(self, request: CreateInboxRuleRequest) -> InboxRule:
        return self.backend.create_inbox_rule(request)

    def update_inbox_rule(self, request: UpdateInboxRuleRequest) -> InboxRule:
        return self.backend.update_inbox_rule(request)

    def delete_inbox_rule(self, request: DeleteInboxRuleRequest) -> ActionResult:
        return self.backend.delete_inbox_rule(request)

    def get_out_of_office(self) -> OutOfOfficeSettings:
        return self._retry_read(self.backend.get_out_of_office)

    def set_out_of_office(self, request: OutOfOfficeSettings) -> OutOfOfficeSettings:
        # A write: never retried, same as every other mutation here.
        return self.backend.set_out_of_office(request)

    def list_emails(self, request: ListEmailsRequest) -> list[EmailSummary]:
        return self._retry_read(lambda: self.backend.list_emails(request))

    def get_email(self, request: GetEmailRequest) -> EmailFull:
        return self._retry_read(lambda: self.backend.get_email(request))

    def get_thread(self, request: GetThreadRequest) -> Thread:
        return self._retry_read(lambda: self.backend.get_thread(request))

    def search_emails(self, request: SearchEmailsRequest) -> list[EmailSummary]:
        return self._retry_read(lambda: self.backend.search_emails(request))

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

    def move_emails(self, request: BulkMoveEmailsRequest) -> BulkResult:
        return self.backend.move_emails(request)

    def copy_emails(self, request: BulkMoveEmailsRequest) -> BulkResult:
        return self.backend.copy_emails(request)

    def delete_emails(self, request: BulkDeleteEmailsRequest) -> BulkResult:
        return self.backend.delete_emails(request)

    def mark_emails(self, request: BulkMarkEmailsRequest) -> BulkResult:
        return self.backend.mark_emails(request)

    def categorize_emails(self, request: BulkCategorizeEmailsRequest) -> BulkResult:
        return self.backend.categorize_emails(request)

    def categorize_email(self, request: CategorizeEmailRequest) -> ActionResult:
        return self.backend.categorize_email(request)

    def list_categories(self, request: ListCategoriesRequest) -> list[CategoryUsage]:
        return self._retry_read(lambda: self.backend.list_categories(request))

    def list_folders(self, request: ListFoldersRequest) -> list[FolderInfo]:
        return self._retry_read(lambda: self.backend.list_folders(request))

    def create_folder(self, request: CreateFolderRequest) -> ActionResult:
        return self.backend.create_folder(request)

    def rename_folder(self, request: RenameFolderRequest) -> ActionResult:
        return self.backend.rename_folder(request)

    def delete_folder(self, request: DeleteFolderRequest) -> ActionResult:
        return self.backend.delete_folder(request)

    def create_draft(self, request: DraftEmailRequest) -> ActionResult:
        return self.backend.create_draft(request)

    def update_draft(self, request: UpdateDraftRequest) -> ActionResult:
        return self.backend.update_draft(request)

    def send_draft(self, request: SendDraftRequest) -> SendResult:
        return self.backend.send_draft(request)

    def add_attachment(self, request: AddAttachmentRequest) -> ActionResult:
        return self.backend.add_attachment(request)

    def delete_attachment(self, request: DeleteAttachmentRequest) -> ActionResult:
        return self.backend.delete_attachment(request)

    def get_attachment(self, request: GetAttachmentRequest) -> AttachmentResult:
        # Not retried despite being read-only: it writes the downloaded content to
        # local disk, so a retried call after a partial download risks leaving an
        # orphaned file behind (_create_new_file never overwrites, it picks a new
        # suffixed name each attempt).
        return self.backend.get_attachment(request)

    def get_email_mime(self, request: GetEmailMimeRequest) -> EmailMimeResult:
        return self._retry_read(lambda: self.backend.get_email_mime(request))

    def list_events(self, request: ListEventsRequest) -> list[CalendarEvent]:
        return self._retry_read(lambda: self.backend.list_events(request))

    def get_event(self, request: GetEventRequest) -> CalendarEvent:
        return self._retry_read(lambda: self.backend.get_event(request))

    def create_event(self, request: CreateEventRequest) -> CreateEventResult:
        return self.backend.create_event(request)

    def update_event(self, request: UpdateEventRequest) -> ActionResult:
        return self.backend.update_event(request)

    def delete_event(self, request: DeleteEventRequest) -> ActionResult:
        return self.backend.delete_event(request)

    def respond_to_invite(self, request: RespondToInviteRequest) -> ActionResult:
        return self.backend.respond_to_invite(request)

    def find_free_slots(self, request: FindFreeSlotsRequest) -> list[FreeSlot]:
        return self._retry_read(lambda: self.backend.find_free_slots(request))

    def delete_events(self, request: BulkDeleteEventsRequest) -> BulkResult:
        return self.backend.delete_events(request)

    def respond_to_invites(self, request: BulkRespondToInvitesRequest) -> BulkResult:
        return self.backend.respond_to_invites(request)

    def get_my_availability(self, request: ListEventsRequest) -> AvailabilityResult:
        return self._retry_read(lambda: self.backend.get_my_availability(request))

    def list_calendars(self) -> list[CalendarInfo]:
        return self._retry_read(self.backend.list_calendars)

    def list_room_lists(self) -> list[RoomListInfo]:
        return self._retry_read(self.backend.list_room_lists)

    def list_rooms(self, request: ListRoomsRequest) -> list[RoomInfo]:
        return self._retry_read(lambda: self.backend.list_rooms(request))

    def search_contacts(self, request: SearchContactsRequest) -> list[ContactSummary]:
        return self._retry_read(lambda: self.backend.search_contacts(request))

    def get_contact(self, request: GetContactRequest) -> ContactFull:
        return self._retry_read(lambda: self.backend.get_contact(request))

    def create_contact(self, request: CreateContactRequest) -> ActionResult:
        return self.backend.create_contact(request)

    def update_contact(self, request: UpdateContactRequest) -> ActionResult:
        return self.backend.update_contact(request)

    def delete_contact(self, request: DeleteContactRequest) -> ActionResult:
        return self.backend.delete_contact(request)
