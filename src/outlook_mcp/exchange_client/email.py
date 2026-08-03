from __future__ import annotations

import tempfile
from datetime import UTC, datetime, timedelta
from pathlib import Path
from typing import Any

from exchangelib import FileAttachment, Folder, HTMLBody, Message
from exchangelib.extended_properties import Flag
from exchangelib.fields import InvalidField

from ..errors import APIError, NotFoundError
from ..models import (
    ActionResult,
    Attachment,
    AttachmentResult,
    CreateFolderRequest,
    DeleteEmailRequest,
    DraftEmailRequest,
    EmailFull,
    EmailSummary,
    FolderActionRequest,
    FolderInfo,
    ForwardEmailRequest,
    GetAttachmentRequest,
    GetEmailRequest,
    ListEmailsRequest,
    ListFoldersRequest,
    MarkEmailRequest,
    ReplyEmailRequest,
    SearchEmailsRequest,
    SendDraftRequest,
    SendEmailRequest,
    SendResult,
)

try:
    Message.get_field_by_fieldname("flag_status")
except InvalidField:
    Message.register("flag_status", Flag)

#: PidTagFlagStatus values: None = not flagged, 1 = completed, 2 = flagged.
_FLAG_STATUS = {"flagged": 2, "complete": 1, "none": None}


class EmailOperationsMixin:
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

    def _preview(self, item: Any) -> str:
        text, _ = self._extract_message_body(item)
        return text[:200]

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
        self._attach_files(message, request.attachments)
        return message

    def _attach_files(self, message: Message, attachments: list[Path]) -> None:
        for path in attachments:
            with Path(path).open("rb") as handle:
                message.attach(FileAttachment(name=Path(path).name, content=handle.read()))

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

    def list_emails(self, request: ListEmailsRequest) -> list[EmailSummary]:
        folder = self._resolve_folder(request.folder)
        qs = folder.all().order_by("-datetime_received")
        filters: dict[str, Any] = {}
        if request.from_address:
            # 'author' is a MailboxField, not an IndexedField, so EWS only supports filtering the field
            # as a whole (exact match against the address) -- '__email_address' is not a valid subfield path.
            filters["author__iexact"] = str(request.from_address)
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
        except Exception as exc:  # noqa: BLE001
            raise self._map_exception(exc) from exc
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
            subject = f"Re: {item.subject or ''}"
            response = (
                item.create_reply_all(subject=subject, body=request.body)
                if request.reply_all
                else item.create_reply(subject=subject, body=request.body)
            )
            return self._send_response_object(response, request.attachments, fallback_id=request.id)
        except Exception as exc:  # noqa: BLE001
            raise self._map_exception(exc, item_id=request.id) from exc

    def forward_email(self, request: ForwardEmailRequest) -> SendResult:
        item = self._fetch_item(request.id)
        try:
            response = item.create_forward(
                subject=f"Fwd: {item.subject or ''}",
                body=request.comment or "",
                to_recipients=[self._mailbox(address) for address in request.to],
            )
            return self._send_response_object(response, request.attachments, fallback_id=request.id)
        except Exception as exc:  # noqa: BLE001
            raise self._map_exception(exc, item_id=request.id) from exc

    def _send_response_object(
        self, response: Any, attachments: list[Path], fallback_id: str
    ) -> SendResult:
        # create_reply/create_reply_all/create_forward response objects have no attachments
        # field of their own, so attachments require saving as a draft first, then attaching.
        if not attachments:
            response.send()
            return SendResult(id=fallback_id, status="sent")
        draft = response.save(self.account.drafts)
        message = self._fetch_item(draft.id)
        self._attach_files(message, attachments)
        message.send()
        return SendResult(id=message.id or fallback_id, status="sent")

    def move_email(self, request: FolderActionRequest) -> ActionResult:
        item = self._fetch_item(request.id)
        destination = self._resolve_folder(request.folder)
        try:
            # item.move() returns None and mutates item.id/changekey in place.
            item.move(to_folder=destination)
            return ActionResult(id=item.id or request.id, status="moved", new_folder=request.folder)
        except Exception as exc:  # noqa: BLE001
            raise self._map_exception(exc, item_id=request.id) from exc

    def copy_email(self, request: FolderActionRequest) -> ActionResult:
        item = self._fetch_item(request.id)
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
        item = self._fetch_item(request.id)
        try:
            if request.hard_delete:
                item.delete()
            else:
                item.move_to_trash()
            return ActionResult(id=request.id, status="deleted")
        except Exception as exc:  # noqa: BLE001
            raise self._map_exception(exc, item_id=request.id) from exc

    def mark_email(self, request: MarkEmailRequest) -> ActionResult:
        item = self._fetch_item(request.id)
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
        if request.flag is not None:
            item.flag_status = _FLAG_STATUS[request.flag]
            updated_fields.append("flag")
            save_fields.append("flag_status")
        try:
            item.save(update_fields=save_fields or None)
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
        max_size_bytes = self.settings.attachment_max_size_mb * 1024 * 1024
        for attachment in getattr(item, "attachments", None) or []:
            attachment_id = getattr(getattr(attachment, "attachment_id", None), "id", None)
            if attachment_id == request.attachment_id:
                declared_size = getattr(attachment, "size", None)
                self._check_attachment_size(declared_size, max_size_bytes)
                filename = self._sanitize_attachment_filename(getattr(attachment, "name", None))
                path = self._unique_path(target_dir / filename)
                content = getattr(attachment, "content", None)
                if content is None:
                    _ = attachment.content
                    content = attachment.content
                self._check_attachment_size(len(content), max_size_bytes)
                path.write_bytes(content)
                return AttachmentResult(
                    filename=filename,
                    size=len(content),
                    saved_path=str(path),
                    content_type=getattr(attachment, "content_type", None),
                )
        raise NotFoundError(request.attachment_id)

    def _check_attachment_size(self, size: int | None, max_size_bytes: int) -> None:
        if size is not None and size > max_size_bytes:
            raise APIError(
                "validation_error",
                "attachment exceeds the configured size limit",
                details=[
                    {
                        "field": "attachment_id",
                        "reason": f"attachment size {size} bytes exceeds "
                        f"ATTACHMENT_MAX_SIZE_MB={self.settings.attachment_max_size_mb}",
                    }
                ],
            )
