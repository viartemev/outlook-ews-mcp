from __future__ import annotations

import base64
import errno
import logging
import os
import re
import stat
import tempfile
from collections import Counter
from datetime import UTC, datetime, timedelta
from pathlib import Path
from typing import Any

from exchangelib import FileAttachment, Folder, HTMLBody, ItemAttachment, Message, Q
from exchangelib.errors import (
    ErrorInvalidRestriction,
    ErrorItemNotFound,
    ErrorUnsupportedPathForQuery,
    ErrorUnsupportedQueryFilter,
)
from exchangelib.extended_properties import ExtendedProperty, Flag
from exchangelib.fields import InvalidField, InvalidFieldForVersion
from exchangelib.folders import FolderCollection
from exchangelib.items import HARD_DELETE, MOVE_TO_DELETED_ITEMS
from exchangelib.properties import ConversationId

from ..errors import APIError, NotFoundError
from ..models import (
    ActionResult,
    Attachment,
    AddAttachmentRequest,
    AttachmentResult,
    BulkCategorizeEmailsRequest,
    BulkDeleteEmailsRequest,
    BulkMarkEmailsRequest,
    DeleteAttachmentRequest,
    BulkItemFailure,
    BulkItemResult,
    BulkMoveEmailsRequest,
    BulkResult,
    CategorizeEmailRequest,
    CategoryUsage,
    CreateFolderRequest,
    DeleteEmailRequest,
    DeleteFolderRequest,
    DraftEmailRequest,
    EmailFull,
    EmailMimeResult,
    EmailSummary,
    FolderActionRequest,
    FolderInfo,
    ForwardEmailRequest,
    GetAttachmentRequest,
    GetEmailMimeRequest,
    GetEmailRequest,
    GetThreadRequest,
    ListCategoriesRequest,
    ListEmailsRequest,
    ListFoldersRequest,
    MarkEmailRequest,
    RenameFolderRequest,
    ReplyEmailRequest,
    SearchEmailsRequest,
    SendDraftRequest,
    SendEmailRequest,
    SendResult,
    Thread,
    UpdateDraftRequest,
)
from .base import BaseEWSBackend

logger = logging.getLogger(__name__)


class FlagStartDate(ExtendedProperty):
    """PidLidTaskStartDate -- the "start" half of an Outlook follow-up flag."""

    distinguished_property_set_id = "Task"
    property_id = 0x8104
    property_type = "SystemTime"


class FlagDueDate(ExtendedProperty):
    """PidLidTaskDueDate -- the "due by" half of an Outlook follow-up flag."""

    distinguished_property_set_id = "Task"
    property_id = 0x8105
    property_type = "SystemTime"


for _field_name, _field_cls in (
    ("flag_status", Flag),
    ("flag_start_date", FlagStartDate),
    ("flag_due_date", FlagDueDate),
):
    try:
        Message.get_field_by_fieldname(_field_name)
    except InvalidField:
        Message.register(_field_name, _field_cls)

#: PidTagFlagStatus values: None = not flagged, 1 = completed, 2 = flagged.
_FLAG_STATUS = {"flagged": 2, "complete": 1, "none": None}

#: Failures that mean this server cannot express a restriction on
#: item:ConversationId at all -- the only case where falling back to a lossy
#: subject match is better than failing. Anything else (throttling, auth, a
#: transient timeout) must surface as itself: a fallback that quietly succeeds
#: would hand back an incomplete thread with no sign the real query never ran.
_UNSUPPORTED_RESTRICTION = (
    ErrorInvalidRestriction,
    ErrorUnsupportedPathForQuery,
    ErrorUnsupportedQueryFilter,
    InvalidFieldForVersion,
)

#: Fields EmailSummary actually reads -- restricting the fetch to these avoids
#: exchangelib issuing an extra GetItem per message and loading full body/
#: attachments/headers just to render a one-line summary.
_EMAIL_SUMMARY_FIELDS = (
    "subject",
    "author",
    "sender",
    "to_recipients",
    "datetime_received",
    "datetime_sent",
    "datetime_created",
    "is_read",
    "has_attachments",
    "text_body",
    "importance",
    "categories",
    "conversation_id",
)

#: Reply/forward markers Outlook writes, in the locales this server sees. The
#: optional bracketed number covers the counted form Exchange sometimes uses
#: ("Re[2]: ").
_REPLY_PREFIX_RE = re.compile(r"^\s*(?:re|ответ|отв)\s*(?:\[\d+\])?\s*:", re.IGNORECASE)
_FORWARD_PREFIX_RE = re.compile(r"^\s*(?:fw|fwd|пересл)\s*(?:\[\d+\])?\s*:", re.IGNORECASE)


def reply_subject(value: str | None) -> str:
    """Prefix with ``Re:`` unless the subject already carries a reply marker."""
    subject = (value or "").strip()
    return subject if _REPLY_PREFIX_RE.match(subject) else f"Re: {subject}".strip()


def forward_subject(value: str | None) -> str:
    """Prefix with ``Fwd:`` unless the subject already carries a forward marker."""
    subject = (value or "").strip()
    return subject if _FORWARD_PREFIX_RE.match(subject) else f"Fwd: {subject}".strip()


def _dedupe_categories(values: list[str]) -> list[str]:
    """Trim and drop case-insensitive duplicates, keeping the first spelling seen."""
    result: list[str] = []
    seen: set[str] = set()
    for value in values:
        name = value.strip()
        key = name.casefold()
        if name and key not in seen:
            seen.add(key)
            result.append(name)
    return result


#: Same markers as above, but matching a whole run of them, so "Re: FW: X"
#: normalises to "X" for thread grouping.
_SUBJECT_PREFIX_RE = re.compile(
    r"^\s*(?:(?:re|fw|fwd|ответ|отв|пересл)\s*(?:\[\d+\])?\s*:\s*)+",
    re.IGNORECASE,
)


def normalize_subject(value: str | None) -> str:
    """Strip reply/forward prefixes so one conversation reads as one subject."""
    return _SUBJECT_PREFIX_RE.sub("", value or "").strip()


class EmailOperationsMixin(BaseEWSBackend):
    def _to_email_summary(self, item: Any) -> EmailSummary:
        return EmailSummary(
            id=item.id,
            subject=item.subject or "",
            **{
                "from": self._email_address(
                    getattr(item, "author", None) or getattr(item, "sender", None)
                )
            },
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
            conversation_id=getattr(getattr(item, "conversation_id", None), "id", None),
        )

    def _attachment_metadata(self, attachment: Any) -> Attachment:
        return Attachment(
            id=getattr(getattr(attachment, "attachment_id", None), "id", None),
            name=getattr(attachment, "name", "attachment"),
            size=getattr(attachment, "size", None),
            content_type=getattr(attachment, "content_type", None),
            downloadable=not isinstance(attachment, ItemAttachment),
        )

    def _to_email_full(self, item: Any) -> EmailFull:
        body_text, body_html = self._extract_message_body(item)
        max_chars = self.settings.email_body_max_chars
        truncated = False
        if len(body_text) > max_chars:
            body_text = body_text[:max_chars]
            truncated = True
        if body_html is not None and len(body_html) > max_chars:
            body_html = body_html[:max_chars]
            truncated = True
        return EmailFull(
            **self._to_email_summary(item).model_dump(by_alias=True),
            cc=self._recipients(getattr(item, "cc_recipients", None)),
            bcc=self._recipients(getattr(item, "bcc_recipients", None)),
            body_text=body_text,
            body_html=body_html,
            attachments=[
                self._attachment_metadata(a) for a in getattr(item, "attachments", None) or []
            ],
            headers=self._headers_to_dict(getattr(item, "headers", None)),
            truncated=truncated,
        )

    def _preview(self, item: Any) -> str:
        text, _ = self._extract_message_body(item)
        return text[:200]

    def _with_signature(self, body: str, body_type: str, include: bool) -> str:
        """Append the configured signature; html bodies get the html signature only,
        text bodies the text one -- no cross-conversion between the two."""
        if not include:
            return body
        signature = (
            self.settings.signature_html if body_type == "html" else self.settings.signature_text
        )
        if not signature:
            return body
        separator = "<br><br>" if body_type == "html" else "\n\n"
        return f"{body}{separator}{signature}"

    def _make_message(self, request: SendEmailRequest | DraftEmailRequest) -> Message:
        body_text = self._with_signature(request.body, request.body_type, request.include_signature)
        body: str | HTMLBody = HTMLBody(body_text) if request.body_type == "html" else body_text
        reply_to: str | None = getattr(request, "reply_to", None)
        importance: str | None = getattr(request, "importance", None)
        message = Message(
            account=self.account,
            folder=self.account.drafts,
            subject=request.subject,
            body=body,
            to_recipients=[self._mailbox(address) for address in request.to],
            cc_recipients=[self._mailbox(address) for address in request.cc],
            bcc_recipients=[self._mailbox(address) for address in request.bcc],
            reply_to=[self._mailbox(reply_to)] if reply_to else None,
            importance=importance.capitalize() if importance else "Normal",
        )
        self._attach_files(message, request.attachments)
        return message

    def _attach_files(self, message: Message, attachments: list[Path]) -> None:
        max_size_bytes = self.settings.attachment_max_size_mb * 1024 * 1024
        max_total_size_bytes = self.settings.attachment_max_total_size_mb * 1024 * 1024
        total_size = 0
        # O_NONBLOCK keeps open() from hanging forever on a FIFO with no writer connected;
        # it's a no-op for the regular files this loop actually accepts (POSIX only applies
        # nonblocking semantics to FIFOs/devices/sockets, never to plain files).
        nonblock_flag = getattr(os, "O_NONBLOCK", 0)
        for path in attachments:
            fd = os.open(path, os.O_RDONLY | nonblock_flag)
            fd_owned = False
            try:
                # Re-validate against the opened fd, not the earlier path-based stat: the
                # file on disk could have been swapped since that check, and a path check
                # alone can't tell a FIFO or /dev/zero from a plain file.
                if not stat.S_ISREG(os.fstat(fd).st_mode):
                    raise APIError(
                        "validation_error",
                        "one or more attachments are not regular files",
                        details=[{"field": "attachments", "reason": f"not a regular file: {path}"}],
                    )
                with os.fdopen(fd, "rb") as handle:
                    fd_owned = True
                    # Bounded to max_size_bytes + 1 so a file that lies about its size
                    # (grown after the earlier stat) can't be read unboundedly into memory.
                    content = handle.read(max_size_bytes + 1)
            finally:
                if not fd_owned:
                    os.close(fd)
            if len(content) > max_size_bytes:
                raise APIError(
                    "validation_error",
                    "one or more attachments exceed the configured size limit",
                    details=[
                        {
                            "field": "attachments",
                            "reason": f"file exceeds EXCHANGE_ATTACHMENT_MAX_SIZE_MB="
                            f"{self.settings.attachment_max_size_mb}: {path}",
                        }
                    ],
                )
            total_size += len(content)
            if total_size > max_total_size_bytes:
                raise APIError(
                    "validation_error",
                    "total attachment size exceeds the configured limit",
                    details=[
                        {
                            "field": "attachments",
                            "reason": f"combined size exceeds EXCHANGE_ATTACHMENT_MAX_TOTAL_SIZE_MB="
                            f"{self.settings.attachment_max_total_size_mb}",
                        }
                    ],
                )
            message.attach(FileAttachment(name=Path(path).name, content=content))

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

    def _sanitize_attachment_filename(self, name: str | None) -> str:
        candidate = Path((name or "attachment.bin").replace("\\", "/")).name.strip()
        if not candidate or candidate in {".", ".."}:
            candidate = "attachment.bin"
        return candidate

    def _create_new_file(self, path: Path) -> tuple[int, Path]:
        """Atomically create a new regular file at path, refusing to follow symlinks.

        Replaces a `path.exists()` check followed by a later `open("wb")`: a symlink
        planted at that name -- pointing at a file that doesn't exist yet -- passes
        `exists()` as False, so the later open would silently follow it and write
        outside EXCHANGE_ATTACHMENT_ROOT. O_EXCL alone already rejects a path that
        names an existing symlink (dangling or not) as EEXIST; O_NOFOLLOW is a
        defense-in-depth backstop against platform differences in that behavior.
        """
        stem = path.stem
        suffix = path.suffix
        candidate = path
        index = 0
        while True:
            try:
                fd = os.open(
                    candidate,
                    os.O_WRONLY | os.O_CREAT | os.O_EXCL | getattr(os, "O_NOFOLLOW", 0),
                    0o600,
                )
                return fd, candidate
            except OSError as exc:
                if exc.errno not in (errno.EEXIST, errno.ELOOP):
                    raise
                index += 1
                candidate = path.with_name(f"{stem}-{index}{suffix}")

    def list_emails(self, request: ListEmailsRequest) -> list[EmailSummary]:
        folder = self._resolve_folder(request.folder)
        qs = folder.all().only(*_EMAIL_SUMMARY_FIELDS).order_by("-datetime_received")
        filters: dict[str, Any] = {}
        if request.from_address:
            # 'author' is a MailboxField, not an IndexedField, so EWS only supports filtering the field
            # as a whole (exact match against the address) -- '__email_address' is not a valid subfield path.
            filters["author__iexact"] = str(request.from_address)
        if request.subject:
            filters["subject__icontains"] = request.subject
        if request.since:
            filters["datetime_received__gte"] = datetime.combine(
                request.since, datetime.min.time(), tzinfo=self.account.default_timezone
            )
        if request.before:
            filters["datetime_received__lt"] = datetime.combine(
                request.before + timedelta(days=1),
                datetime.min.time(),
                tzinfo=self.account.default_timezone,
            )
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
        item = self._fetch_item(request.id, expected_type=Message)
        return self._to_email_full(item)

    #: Standard EWS folder class for mail folders; excludes calendar/contacts/tasks
    #: folders when a search walks the whole mailbox.
    _MAIL_FOLDER_CLASS = "IPF.Note"

    def get_thread(self, request: GetThreadRequest) -> Thread:
        conversation_id = request.conversation_id
        subject: str | None = None
        anchor: Any = None
        if request.id:
            anchor = self._fetch_item(request.id, expected_type=Message)
            conversation_id = getattr(getattr(anchor, "conversation_id", None), "id", None)
            subject = anchor.subject or ""

        collected: dict[str, Any] = {}
        if anchor is not None and getattr(anchor, "id", None):
            # The anchor may sit in a folder outside `folders`; an empty thread for
            # a message we are holding would be a wrong answer, not a partial one.
            collected[anchor.id] = anchor
        for name in request.folders:
            folder = self._resolve_folder(name)
            for item in self._thread_items(folder, conversation_id, subject, request.limit):
                item_id = getattr(item, "id", None)
                # The same message surfaces twice when folders overlap.
                if item_id and item_id not in collected:
                    collected[item_id] = item

        found = sorted(collected.values(), key=self._received_at)
        truncated = len(found) > request.limit
        messages = [self._to_email_full(item) for item in found[-request.limit :]]
        return Thread(
            conversation_id=conversation_id,
            subject=normalize_subject(subject or (messages[0].subject if messages else "")),
            message_count=len(messages),
            messages=messages,
            truncated=truncated,
        )

    def _thread_items(
        self, folder: Folder, conversation_id: str | None, subject: str | None, limit: int
    ) -> list[Any]:
        if conversation_id:
            try:
                qs = folder.filter(conversation_id=ConversationId(id=conversation_id))
                return list(qs.order_by("-datetime_received")[:limit])
            except _UNSUPPORTED_RESTRICTION as exc:
                if subject is None:
                    # Nothing to fall back to. Reporting an empty conversation for a
                    # query we could not run would be a wrong answer, not a partial one.
                    raise self._map_exception(exc) from exc
                logger.info(
                    "server does not support restricting on item:ConversationId (%s); "
                    "falling back to a subject match",
                    type(exc).__name__,
                )
            except Exception as exc:  # noqa: BLE001
                # Anything else -- throttling, auth, a transient timeout -- says
                # nothing about whether the restriction is supported. Surface it
                # like the rest of the module does instead of quietly returning a
                # lossy thread that hides the real failure.
                raise self._map_exception(exc) from exc

        normalized = normalize_subject(subject)
        if not normalized:
            return []
        try:
            qs = folder.filter(subject__icontains=normalized).order_by("-datetime_received")
            items = list(qs[:limit])
        except Exception as exc:  # noqa: BLE001
            raise self._map_exception(exc) from exc
        # icontains also matches subjects that merely embed this one.
        target = normalized.casefold()
        return [item for item in items if normalize_subject(item.subject).casefold() == target]

    def _received_at(self, item: Any) -> datetime:
        return (
            getattr(item, "datetime_received", None)
            or getattr(item, "datetime_sent", None)
            or getattr(item, "datetime_created", None)
            or datetime.now(UTC)
        )

    def search_emails(self, request: SearchEmailsRequest) -> list[EmailSummary]:
        searchable: Folder | FolderCollection
        try:
            if request.folder:
                searchable = self._resolve_folder(request.folder)
            else:
                # See list_calendars() for why walk() is fine here: with no folder given we
                # must discover every mail folder, which has no id to feed _get_folder_by_id,
                # and walk() is a single cached Deep-traversal call, not one call per folder.
                searchable = FolderCollection(
                    account=self.account,
                    folders=[
                        folder
                        for folder in self.account.root.walk()
                        if getattr(folder, "folder_class", None) == self._MAIL_FOLDER_CLASS
                    ],
                )
            if request.aqs:
                # A QueryString is sent to EWS as its own FindItem element and cannot
                # be combined with a Restriction, so no other filter is added here.
                items = self._search_items(searchable, Q(request.aqs), request.limit)
            else:
                query = (
                    Q(subject__icontains=request.query)
                    | Q(text_body__icontains=request.query)
                    | Q(author__icontains=request.query)
                )
                try:
                    items = self._search_items(searchable, query, request.limit)
                except ErrorUnsupportedPathForQuery:
                    # Exchange 2019 on-prem rejects a substring restriction on
                    # item:TextBody outright ("The property can not be used with
                    # this type of restriction"), killing the whole search. Seen
                    # live; subject and sender still match, so retry without the
                    # body leg rather than failing the call.
                    logger.info(
                        "this server rejects substring matching on the body; "
                        "searching subject and sender only"
                    )
                    query = Q(subject__icontains=request.query) | Q(author__icontains=request.query)
                    items = self._search_items(searchable, query, request.limit)
        except Exception as exc:  # noqa: BLE001
            raise self._map_exception(exc) from exc
        return [self._to_email_summary(item) for item in items]

    def _search_items(
        self, searchable: Folder | FolderCollection, query: Q, limit: int
    ) -> list[Any]:
        qs = searchable.filter(query).only(*_EMAIL_SUMMARY_FIELDS).order_by("-datetime_received")
        return list(qs[:limit])

    def send_email(self, request: SendEmailRequest) -> SendResult:
        message = self._make_message(request)
        try:
            message.send_and_save()
            # send_and_save() doesn't hand back the id of the sent-and-saved copy.
            return SendResult(id=message.id or None, status="sent")
        except Exception as exc:  # noqa: BLE001
            raise self._map_exception(exc) from exc

    def reply_email(self, request: ReplyEmailRequest) -> SendResult:
        item = self._fetch_item(request.id, expected_type=Message)
        try:
            subject = reply_subject(item.subject)
            body = self._with_signature(request.body, "text", request.include_signature)
            response = (
                item.create_reply_all(subject=subject, body=body)
                if request.reply_all
                else item.create_reply(subject=subject, body=body)
            )
            return self._send_response_object(response, request.attachments)
        except Exception as exc:  # noqa: BLE001
            raise self._map_exception(exc, item_id=request.id) from exc

    def forward_email(self, request: ForwardEmailRequest) -> SendResult:
        item = self._fetch_item(request.id, expected_type=Message)
        try:
            response = item.create_forward(
                subject=forward_subject(item.subject),
                body=self._with_signature(request.comment or "", "text", request.include_signature),
                to_recipients=[self._mailbox(address) for address in request.to],
            )
            return self._send_response_object(response, request.attachments)
        except Exception as exc:  # noqa: BLE001
            raise self._map_exception(exc, item_id=request.id) from exc

    def _send_response_object(self, response: Any, attachments: list[Path]) -> SendResult:
        # create_reply/create_reply_all/create_forward response objects have no attachments
        # field of their own, so attachments require saving as a draft first, then attaching.
        # Neither path hands back a trustworthy id for the sent message: response.send()
        # doesn't save/return the sent item at all, and the draft's id stops being valid
        # the moment message.send() moves it out of Drafts -- so no id is fabricated here.
        if not attachments:
            response.send()
            return SendResult(id=None, status="sent")
        draft = response.save(self.account.drafts)
        message = self._fetch_item(draft.id, folder=self.account.drafts, expected_type=Message)
        self._attach_files(message, attachments)
        message.send()
        return SendResult(id=None, status="sent")

    def move_email(self, request: FolderActionRequest) -> ActionResult:
        item = self._fetch_item(request.id, expected_type=Message)
        destination = self._resolve_folder(request.folder)
        try:
            # item.move() returns None and mutates item.id/changekey in place.
            item.move(to_folder=destination)
            return ActionResult(id=item.id or request.id, status="moved", new_folder=request.folder)
        except Exception as exc:  # noqa: BLE001
            raise self._map_exception(exc, item_id=request.id) from exc

    def copy_email(self, request: FolderActionRequest) -> ActionResult:
        item = self._fetch_item(request.id, expected_type=Message)
        destination = self._resolve_folder(request.folder)
        try:
            # item.copy() returns an (id, changekey) tuple for the new item, unlike item.move().
            result = item.copy(to_folder=destination)
            new_id = result[0] if result else None
            return ActionResult(
                id=request.id,
                status="copied",
                new_folder=request.folder,
                new_id=new_id,
            )
        except Exception as exc:  # noqa: BLE001
            raise self._map_exception(exc, item_id=request.id) from exc

    def delete_email(self, request: DeleteEmailRequest) -> ActionResult:
        item = self._fetch_item(request.id, expected_type=Message)
        try:
            if request.hard_delete:
                item.delete()
            else:
                item.move_to_trash()
            return ActionResult(id=request.id, status="deleted")
        except Exception as exc:  # noqa: BLE001
            raise self._map_exception(exc, item_id=request.id) from exc

    def mark_email(self, request: MarkEmailRequest) -> ActionResult:
        item = self._fetch_item(request.id, expected_type=Message)
        updated_fields: list[str] = []
        save_fields: list[str] = []
        if request.read is not None:
            item.is_read = request.read
            updated_fields.append("read")
            save_fields.append("is_read")
        if request.importance is not None:
            item.importance = request.importance.capitalize()
            updated_fields.append("importance")
            save_fields.append("importance")
        flag = request.flag
        if flag is None and (
            request.flag_start_date is not None or request.flag_due_date is not None
        ):
            # Outlook renders nothing for a due date on an unflagged message, so
            # dates alone mean "flag it, due then".
            flag = "flagged"
        if flag is not None:
            item.flag_status = _FLAG_STATUS[flag]
            updated_fields.append("flag")
            save_fields.append("flag_status")
        if request.flag_start_date is not None:
            item.flag_start_date = self._to_ews_datetime(request.flag_start_date)
            updated_fields.append("flag_start_date")
            save_fields.append("flag_start_date")
        if request.flag_due_date is not None:
            item.flag_due_date = self._to_ews_datetime(request.flag_due_date)
            updated_fields.append("flag_due_date")
            save_fields.append("flag_due_date")
        if not save_fields:
            # Nothing to change -- saving anyway would still write every loaded field back.
            return ActionResult(id=request.id, status="updated", updated_fields=updated_fields)
        try:
            item.save(update_fields=save_fields)
            return ActionResult(id=request.id, status="updated", updated_fields=updated_fields)
        except Exception as exc:  # noqa: BLE001
            raise self._map_exception(exc, item_id=request.id) from exc

    def _bulk_result(self, ids: list[str], results: list[Any]) -> BulkResult:
        if len(results) != len(ids):
            # Moving/copying into a public folder or another mailbox returns no
            # per-item ids at all; pad so every input id still gets a verdict.
            results = list(results) + [None] * (len(ids) - len(results))
        succeeded: list[BulkItemResult] = []
        failed: list[BulkItemFailure] = []
        for item_id, result in zip(ids, results):
            if isinstance(result, Exception):
                mapped = self._map_exception(result, item_id=item_id)
                failed.append(
                    BulkItemFailure(id=item_id, error=mapped.code, message=mapped.message)
                )
            else:
                new_id = result[0] if isinstance(result, tuple) else None
                succeeded.append(BulkItemResult(id=item_id, new_id=new_id))
        return BulkResult(succeeded=succeeded, failed=failed)

    def move_emails(self, request: BulkMoveEmailsRequest) -> BulkResult:
        destination = self._resolve_folder(request.folder)
        ids = [(item_id, None) for item_id in request.ids]
        try:
            results = self.account.bulk_move(ids=ids, to_folder=destination)
        except Exception as exc:  # noqa: BLE001
            raise self._map_exception(exc) from exc
        return self._bulk_result(request.ids, results)

    def copy_emails(self, request: BulkMoveEmailsRequest) -> BulkResult:
        destination = self._resolve_folder(request.folder)
        ids = [(item_id, None) for item_id in request.ids]
        try:
            results = self.account.bulk_copy(ids=ids, to_folder=destination)
        except Exception as exc:  # noqa: BLE001
            raise self._map_exception(exc) from exc
        return self._bulk_result(request.ids, results)

    def delete_emails(self, request: BulkDeleteEmailsRequest) -> BulkResult:
        delete_type = HARD_DELETE if request.hard_delete else MOVE_TO_DELETED_ITEMS
        ids = [(item_id, None) for item_id in request.ids]
        try:
            results = self.account.bulk_delete(ids=ids, delete_type=delete_type)
        except Exception as exc:  # noqa: BLE001
            raise self._map_exception(exc) from exc
        return self._bulk_result(request.ids, results)

    def mark_emails(self, request: BulkMarkEmailsRequest) -> BulkResult:
        # Not a native EWS batch operation the way move/copy/delete are, so this
        # loops through mark_email() per id via the shared _bulk() helper.
        return self._bulk(
            request.ids,
            lambda item_id: self.mark_email(
                MarkEmailRequest(
                    id=item_id,
                    read=request.read,
                    flag=request.flag,
                    importance=request.importance,
                    flag_start_date=request.flag_start_date,
                    flag_due_date=request.flag_due_date,
                )
            ),
        )

    def categorize_emails(self, request: BulkCategorizeEmailsRequest) -> BulkResult:
        return self._bulk(
            request.ids,
            lambda item_id: self.categorize_email(
                CategorizeEmailRequest(id=item_id, categories=request.categories, mode=request.mode)
            ),
        )

    def categorize_email(self, request: CategorizeEmailRequest) -> ActionResult:
        item = self._fetch_item(request.id, expected_type=Message)
        current = list(getattr(item, "categories", None) or [])
        if request.mode == "set":
            updated = _dedupe_categories(request.categories)
        elif request.mode == "add":
            updated = _dedupe_categories([*current, *request.categories])
        else:
            removed = {name.strip().casefold() for name in request.categories}
            updated = [name for name in current if name.strip().casefold() not in removed]
        if updated == current:
            return ActionResult(
                id=request.id, status="updated", updated_fields=[], categories=updated
            )
        # Exchange clears the property on None; an empty list is not the same thing.
        item.categories = updated or None
        try:
            item.save(update_fields=["categories"])
        except Exception as exc:  # noqa: BLE001
            raise self._map_exception(exc, item_id=request.id) from exc
        return ActionResult(
            id=request.id,
            status="updated",
            updated_fields=["categories"],
            categories=updated,
        )

    #: Outlook's 25 preset category colours, by the index stored in the
    #: CategoryList XML. -1 (or anything unknown) means "no colour".
    _CATEGORY_COLORS = (
        "red",
        "orange",
        "peach",
        "yellow",
        "green",
        "teal",
        "olive",
        "blue",
        "purple",
        "maroon",
        "steel",
        "dark_steel",
        "gray",
        "dark_gray",
        "black",
        "dark_red",
        "dark_orange",
        "brown",
        "dark_yellow",
        "dark_green",
        "dark_teal",
        "dark_olive",
        "dark_blue",
        "dark_purple",
        "dark_maroon",
    )

    def _master_category_list(self) -> list[CategoryUsage] | None:
        """Read the mailbox master category list, or None when it does not exist.

        Outlook stores it as the "CategoryList" user configuration on the
        Calendar folder -- an XML document, not a folder of items. This is the
        list the Outlook UI shows, including categories that are defined but not
        currently applied to any message.
        """
        import xml.etree.ElementTree as ET

        try:
            config = self.account.calendar.get_user_configuration(name="CategoryList")
        except ErrorItemNotFound:
            return None
        except Exception as exc:  # noqa: BLE001
            raise self._map_exception(exc) from exc
        xml_data = getattr(config, "xml_data", None)
        if not xml_data:
            return None
        try:
            root = ET.fromstring(xml_data)
        except ET.ParseError:
            logger.warning("CategoryList user configuration holds unparseable XML; ignoring it")
            return None
        namespace = root.tag.partition("}")[0].lstrip("{") if "}" in root.tag else ""
        tag = f"{{{namespace}}}category" if namespace else "category"
        result = []
        for category in root.iter(tag):
            name = (category.get("name") or "").strip()
            if not name:
                continue
            try:
                color_index = int(category.get("color", "-1"))
            except ValueError:
                color_index = -1
            color = (
                self._CATEGORY_COLORS[color_index]
                if 0 <= color_index < len(self._CATEGORY_COLORS)
                else None
            )
            try:
                count = max(0, int(category.get("usageCount", "0")))
            except ValueError:
                count = 0
            result.append(CategoryUsage(name=name, count=count, color=color))
        return result or None

    def list_categories(self, request: ListCategoriesRequest) -> list[CategoryUsage]:
        master = self._master_category_list()
        if master is not None:
            return sorted(master, key=lambda usage: (-usage.count, usage.name.casefold()))
        # No master list on this mailbox -- fall back to counting what the most
        # recent messages actually carry. Colours are unknown on this path.
        counts: Counter[str] = Counter()
        labels: dict[str, str] = {}
        for name in request.folders:
            folder = self._resolve_folder(name)
            qs = folder.all().only("categories", "datetime_received")
            try:
                items = list(qs.order_by("-datetime_received")[: request.limit])
            except Exception as exc:  # noqa: BLE001
                raise self._map_exception(exc) from exc
            for item in items:
                for category in getattr(item, "categories", None) or []:
                    label = category.strip()
                    if not label:
                        continue
                    key = label.casefold()
                    labels.setdefault(key, label)
                    counts[key] += 1
        return [CategoryUsage(name=labels[key], count=count) for key, count in counts.most_common()]

    def list_folders(self, request: ListFoldersRequest) -> list[FolderInfo]:
        folder = self._resolve_folder(request.parent)
        try:
            return (
                [self._to_folder_info(child, request.depth - 1) for child in folder.children]
                if request.depth != 0
                else [self._to_folder_info(folder, 0)]
            )
        except Exception as exc:  # noqa: BLE001
            raise self._map_exception(exc) from exc

    def create_folder(self, request: CreateFolderRequest) -> ActionResult:
        parent = self._resolve_folder(request.parent)
        folder = Folder(parent=parent, name=request.name)
        try:
            folder.save()
            return ActionResult(
                id=getattr(folder, "id", ""), status="created", path=self._folder_path(folder)
            )
        except Exception as exc:  # noqa: BLE001
            raise self._map_exception(exc) from exc

    def _ensure_not_distinguished(self, folder: Folder, *, action: str) -> None:
        # Exchange refuses server-side too (ErrorDeleteDistinguishedFolder), but a
        # client-side check gives a clearer message and also covers rename, which
        # the server otherwise allows even for folders like Inbox or Calendar.
        if getattr(folder, "is_distinguished", False):
            raise APIError(
                "validation_error",
                f"cannot {action} a distinguished folder (e.g. Inbox, Sent Items, Calendar)",
                details=[{"field": "folder", "reason": "folder is distinguished"}],
            )

    def rename_folder(self, request: RenameFolderRequest) -> ActionResult:
        folder = self._resolve_folder(request.folder)
        self._ensure_not_distinguished(folder, action="rename")
        folder.name = request.name
        try:
            folder.save(update_fields=["name"])
            return ActionResult(
                id=getattr(folder, "id", ""), status="renamed", path=self._folder_path(folder)
            )
        except Exception as exc:  # noqa: BLE001
            raise self._map_exception(exc) from exc

    def delete_folder(self, request: DeleteFolderRequest) -> ActionResult:
        folder = self._resolve_folder(request.folder)
        self._ensure_not_distinguished(folder, action="delete")
        folder_id = getattr(folder, "id", "")
        try:
            folder.delete(delete_type=HARD_DELETE if request.hard_delete else MOVE_TO_DELETED_ITEMS)
            return ActionResult(id=folder_id, status="deleted")
        except Exception as exc:  # noqa: BLE001
            raise self._map_exception(exc) from exc

    def create_draft(self, request: DraftEmailRequest) -> ActionResult:
        message = self._make_message(request)
        try:
            message.save()
            return ActionResult(id=message.id or "", status="draft")
        except Exception as exc:  # noqa: BLE001
            raise self._map_exception(exc) from exc

    def update_draft(self, request: UpdateDraftRequest) -> ActionResult:
        item = self._fetch_item(request.id, folder=self.account.drafts, expected_type=Message)
        updated_fields: list[str] = []
        save_fields: list[str] = []
        # A field name in `fields_set` means the caller explicitly included it in the
        # request (even as an empty list) -- an omitted field is left untouched. Mirrors
        # the fields_set handling in update_contact/update_event.
        fields_set = request.model_fields_set
        field_map = {
            "to": "to_recipients",
            "subject": "subject",
            "body": "body",
            "cc": "cc_recipients",
            "bcc": "bcc_recipients",
        }
        for request_field, item_field in field_map.items():
            if request_field not in fields_set:
                continue
            value = getattr(request, request_field)
            if request_field == "body" and value is not None:
                value = HTMLBody(value) if request.body_type == "html" else value
            elif request_field in ("to", "cc", "bcc"):
                value = [self._mailbox(address) for address in value or []]
            setattr(item, item_field, value)
            updated_fields.append(request_field)
            save_fields.append(item_field)
        try:
            if "attachments" in fields_set:
                item.detach(list(item.attachments))
                self._attach_files(item, request.attachments or [])
                updated_fields.append("attachments")
            if save_fields:
                item.save(update_fields=save_fields)
            return ActionResult(id=request.id, status="updated", updated_fields=updated_fields)
        except Exception as exc:  # noqa: BLE001
            raise self._map_exception(exc, item_id=request.id) from exc

    def send_draft(self, request: SendDraftRequest) -> SendResult:
        item = self._fetch_item(request.id, folder=self.account.drafts, expected_type=Message)
        try:
            item.send_and_save()
            # The draft's own id stops being valid the moment it's sent (it moves out
            # of Drafts), so request.id must not be echoed back as if still usable.
            return SendResult(id=None, status="sent")
        except Exception as exc:  # noqa: BLE001
            raise self._map_exception(exc, item_id=request.id) from exc

    def add_attachment(self, request: AddAttachmentRequest) -> ActionResult:
        path = request.path
        if not path.is_absolute():
            root = self.settings.attachment_root
            if root is None:
                raise APIError(
                    "validation_error",
                    "relative attachment paths need EXCHANGE_ATTACHMENT_ROOT",
                    details=[
                        {
                            "field": "path",
                            "reason": "set EXCHANGE_ATTACHMENT_ROOT or pass an absolute path",
                        }
                    ],
                )
            # A relative path means "relative to the configured root" -- resolving
            # it against the server's cwd would attach whatever happens to lie there.
            path = root / path
        if not path.is_file():
            raise APIError(
                "validation_error",
                "attachment file does not exist",
                details=[{"field": "path", "reason": f"no such file: {path}"}],
            )
        item = self._fetch_item(request.email_id, expected_type=Message)
        try:
            # Same fd-validated read path as outgoing mail, so a FIFO or a file
            # swapped after the stat gets rejected here too.
            self._attach_files(item, [path])
        except APIError:
            raise
        except Exception as exc:  # noqa: BLE001
            raise self._map_exception(exc, item_id=request.email_id) from exc
        return ActionResult(id=request.email_id, status="updated", updated_fields=["attachments"])

    def delete_attachment(self, request: DeleteAttachmentRequest) -> ActionResult:
        item = self._fetch_item(request.email_id, expected_type=Message)
        for attachment in getattr(item, "attachments", None) or []:
            attachment_id = getattr(getattr(attachment, "attachment_id", None), "id", None)
            if attachment_id == request.attachment_id:
                try:
                    item.detach(attachment)
                except Exception as exc:  # noqa: BLE001
                    raise self._map_exception(exc, item_id=request.email_id) from exc
                return ActionResult(
                    id=request.email_id, status="updated", updated_fields=["attachments"]
                )
        raise NotFoundError(request.attachment_id)

    def get_email_mime(self, request: GetEmailMimeRequest) -> EmailMimeResult:
        item = self._fetch_item(request.id, expected_type=Message)
        try:
            raw = item.mime_content
        except Exception as exc:  # noqa: BLE001
            raise self._map_exception(exc, item_id=request.id) from exc
        if raw is None:
            raise APIError(
                "exchange_error", "exchange did not return MIME content for this message"
            )
        if isinstance(raw, str):
            raw = raw.encode("utf-8")
        filename = self._sanitize_attachment_filename(f"{item.subject or 'message'}.eml")
        return EmailMimeResult(
            id=request.id,
            filename=filename,
            size=len(raw),
            mime_base64=base64.b64encode(raw).decode("ascii"),
        )

    def get_attachment(self, request: GetAttachmentRequest) -> AttachmentResult:
        item = self._fetch_item(request.email_id, expected_type=Message)
        target_dir = Path(request.save_path) if request.save_path else Path(tempfile.gettempdir())
        target_dir.mkdir(parents=True, exist_ok=True)
        max_size_bytes = self.settings.attachment_max_size_mb * 1024 * 1024
        for attachment in getattr(item, "attachments", None) or []:
            attachment_id = getattr(getattr(attachment, "attachment_id", None), "id", None)
            if attachment_id == request.attachment_id:
                if isinstance(attachment, ItemAttachment):
                    raise APIError(
                        "validation_error",
                        "attachment is an embedded Exchange item and cannot be "
                        "downloaded as a file",
                        details=[
                            {
                                "field": "attachment_id",
                                "reason": "embedded item attachments (email/calendar/"
                                "contact items) are not supported by get_attachment",
                            }
                        ],
                    )
                self._check_attachment_size(getattr(attachment, "size", None), max_size_bytes)
                filename = self._sanitize_attachment_filename(getattr(attachment, "name", None))
                fd, path = self._create_new_file(target_dir / filename)
                try:
                    size = self._save_attachment(attachment, fd, max_size_bytes)
                except APIError:
                    path.unlink(missing_ok=True)
                    raise
                except Exception as exc:  # noqa: BLE001
                    path.unlink(missing_ok=True)
                    raise self._map_exception(exc, item_id=request.email_id) from exc
                return AttachmentResult(
                    filename=filename,
                    size=size,
                    saved_path=str(path),
                    content_type=getattr(attachment, "content_type", None),
                )
        raise NotFoundError(request.attachment_id)

    def _save_attachment(self, attachment: Any, fd: int, max_size_bytes: int) -> int:
        with os.fdopen(fd, "wb") as dest:
            if not hasattr(attachment, "fp"):
                content = attachment.content
                self._check_attachment_size(len(content), max_size_bytes)
                dest.write(content)
                return len(content)
            # FileAttachment.fp streams the content from the GetAttachment service in
            # chunks instead of buffering the whole thing in memory (as .content does),
            # so an oversized attachment can be caught -- and the partial file
            # discarded by the caller -- without ever holding it all in RAM.
            written = 0
            with attachment.fp as source:
                while True:
                    chunk = source.read(65536)
                    if not chunk:
                        break
                    written += len(chunk)
                    self._check_attachment_size(written, max_size_bytes)
                    dest.write(chunk)
            return written

    def _check_attachment_size(self, size: int | None, max_size_bytes: int) -> None:
        if size is not None and size > max_size_bytes:
            raise APIError(
                "validation_error",
                "attachment exceeds the configured size limit",
                details=[
                    {
                        "field": "attachment_id",
                        "reason": f"attachment size {size} bytes exceeds "
                        f"EXCHANGE_ATTACHMENT_MAX_SIZE_MB={self.settings.attachment_max_size_mb}",
                    }
                ],
            )
