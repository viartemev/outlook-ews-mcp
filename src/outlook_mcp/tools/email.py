from __future__ import annotations

from pathlib import Path

from pydantic import ValidationError

from ..config import Settings
from ..errors import APIError, validation_error_from_pydantic
from ..exchange_client import ExchangeClient
from ..models import (
    CreateFolderRequest,
    DeleteEmailRequest,
    DraftEmailRequest,
    FolderActionRequest,
    GetAttachmentRequest,
    GetEmailRequest,
    ListEmailsRequest,
    ListFoldersRequest,
    MarkEmailRequest,
    ReplyEmailRequest,
    SearchEmailsRequest,
    SendDraftRequest,
    SendEmailRequest,
    ForwardEmailRequest,
    dump_model,
)


def list_emails(client: ExchangeClient, arguments: dict) -> list[dict]:
    try:
        request = ListEmailsRequest.model_validate(arguments)
    except ValidationError as exc:
        raise validation_error_from_pydantic(exc) from exc
    return dump_model(client.list_emails(request))


def get_email(client: ExchangeClient, arguments: dict) -> dict:
    try:
        request = GetEmailRequest.model_validate(arguments)
    except ValidationError as exc:
        raise validation_error_from_pydantic(exc) from exc
    return dump_model(client.get_email(request))


def search_emails(client: ExchangeClient, arguments: dict) -> list[dict]:
    try:
        request = SearchEmailsRequest.model_validate(arguments)
    except ValidationError as exc:
        raise validation_error_from_pydantic(exc) from exc
    return dump_model(client.search_emails(request))


def send_email(client: ExchangeClient, arguments: dict) -> dict:
    try:
        request = SendEmailRequest.model_validate(arguments)
    except ValidationError as exc:
        raise validation_error_from_pydantic(exc) from exc

    _validate_attachments(request.attachments, client.settings)
    return dump_model(client.send_email(request))


def reply_email(client: ExchangeClient, arguments: dict) -> dict:
    try:
        request = ReplyEmailRequest.model_validate(arguments)
    except ValidationError as exc:
        raise validation_error_from_pydantic(exc) from exc
    _validate_attachments(request.attachments, client.settings)
    return dump_model(client.reply_email(request))


def forward_email(client: ExchangeClient, arguments: dict) -> dict:
    try:
        request = ForwardEmailRequest.model_validate(arguments)
    except ValidationError as exc:
        raise validation_error_from_pydantic(exc) from exc
    _validate_attachments(request.attachments, client.settings)
    return dump_model(client.forward_email(request))


def move_email(client: ExchangeClient, arguments: dict) -> dict:
    try:
        request = FolderActionRequest.model_validate(arguments)
    except ValidationError as exc:
        raise validation_error_from_pydantic(exc) from exc
    return dump_model(client.move_email(request))


def copy_email(client: ExchangeClient, arguments: dict) -> dict:
    try:
        request = FolderActionRequest.model_validate(arguments)
    except ValidationError as exc:
        raise validation_error_from_pydantic(exc) from exc
    return dump_model(client.copy_email(request))


def delete_email(client: ExchangeClient, arguments: dict) -> dict:
    try:
        request = DeleteEmailRequest.model_validate(arguments)
    except ValidationError as exc:
        raise validation_error_from_pydantic(exc) from exc
    return dump_model(client.delete_email(request))


def mark_email(client: ExchangeClient, arguments: dict) -> dict:
    try:
        request = MarkEmailRequest.model_validate(arguments)
    except ValidationError as exc:
        raise validation_error_from_pydantic(exc) from exc
    return dump_model(client.mark_email(request))


def list_folders(client: ExchangeClient, arguments: dict) -> list[dict]:
    try:
        request = ListFoldersRequest.model_validate(arguments)
    except ValidationError as exc:
        raise validation_error_from_pydantic(exc) from exc
    return dump_model(client.list_folders(request))


def create_folder(client: ExchangeClient, arguments: dict) -> dict:
    try:
        request = CreateFolderRequest.model_validate(arguments)
    except ValidationError as exc:
        raise validation_error_from_pydantic(exc) from exc
    return dump_model(client.create_folder(request))


def create_draft(client: ExchangeClient, arguments: dict) -> dict:
    try:
        request = DraftEmailRequest.model_validate(arguments)
    except ValidationError as exc:
        raise validation_error_from_pydantic(exc) from exc
    _validate_attachments(request.attachments, client.settings)
    return dump_model(client.create_draft(request))


def send_draft(client: ExchangeClient, arguments: dict) -> dict:
    try:
        request = SendDraftRequest.model_validate(arguments)
    except ValidationError as exc:
        raise validation_error_from_pydantic(exc) from exc
    return dump_model(client.send_draft(request))


def get_attachment(client: ExchangeClient, arguments: dict) -> dict:
    try:
        request = GetAttachmentRequest.model_validate(arguments)
    except ValidationError as exc:
        raise validation_error_from_pydantic(exc) from exc
    if request.save_path is not None:
        _validate_within_root([request.save_path], client.settings.attachment_root, "save_path")
    return dump_model(client.get_attachment(request))


def _validate_within_root(paths: list[Path], root: Path | None, field: str) -> None:
    if root is None:
        return
    resolved_root = root.resolve()
    outside = [
        str(path) for path in paths if not Path(path).resolve().is_relative_to(resolved_root)
    ]
    if outside:
        raise APIError(
            "validation_error",
            f"one or more {field} paths are outside the configured EXCHANGE_ATTACHMENT_ROOT",
            details=[
                {"field": field, "reason": f"path outside EXCHANGE_ATTACHMENT_ROOT: {path}"}
                for path in outside
            ],
        )


def _validate_attachments(attachments: list[Path], settings: Settings) -> None:
    _validate_within_root(attachments, settings.attachment_root, "attachments")

    if len(attachments) > settings.attachment_max_count:
        raise APIError(
            "validation_error",
            "too many attachments",
            details=[
                {
                    "field": "attachments",
                    "reason": f"{len(attachments)} attachments exceed "
                    f"ATTACHMENT_MAX_COUNT={settings.attachment_max_count}",
                }
            ],
        )

    missing = [str(path) for path in attachments if not path.exists()]
    if missing:
        raise APIError(
            "validation_error",
            "one or more attachments do not exist",
            details=[
                {"field": "attachments", "reason": f"missing file: {path}"} for path in missing
            ],
        )

    max_size_bytes = settings.attachment_max_size_mb * 1024 * 1024
    sizes = [(path, path.stat().st_size) for path in attachments]
    oversized = [{"path": str(path), "size": size} for path, size in sizes if size > max_size_bytes]
    if oversized:
        raise APIError(
            "validation_error",
            "one or more attachments exceed the configured size limit",
            details=[
                {
                    "field": "attachments",
                    "reason": f"file exceeds ATTACHMENT_MAX_SIZE_MB="
                    f"{settings.attachment_max_size_mb}: {item['path']} ({item['size']} bytes)",
                }
                for item in oversized
            ],
        )

    max_total_size_bytes = settings.attachment_max_total_size_mb * 1024 * 1024
    total_size = sum(size for _, size in sizes)
    if total_size > max_total_size_bytes:
        raise APIError(
            "validation_error",
            "total attachment size exceeds the configured limit",
            details=[
                {
                    "field": "attachments",
                    "reason": f"combined size {total_size} bytes exceeds "
                    f"ATTACHMENT_MAX_TOTAL_SIZE_MB={settings.attachment_max_total_size_mb}",
                }
            ],
        )
