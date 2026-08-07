from __future__ import annotations

from pathlib import Path
from typing import Any

from ..config import Settings
from ..errors import APIError
from ..exchange_client import ExchangeClient
from ..models import (
    CategorizeEmailRequest,
    CreateFolderRequest,
    DeleteEmailRequest,
    DeleteFolderRequest,
    DraftEmailRequest,
    FolderActionRequest,
    GetAttachmentRequest,
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
    ForwardEmailRequest,
    UpdateDraftRequest,
)
from .common import tool_handler


# request is whichever subtype tool_handler validated (DraftEmailRequest /
# SendEmailRequest / ReplyEmailRequest / ForwardEmailRequest, or
# GetAttachmentRequest below) -- narrower than the shared ExchangeModel base
# these hooks are declared against in RequestHook, so it's typed as Any here.
def _validate_outgoing_attachments(client: ExchangeClient, request: Any) -> None:
    _validate_attachments(request.attachments, client.settings)


def _validate_outgoing_attachments_if_set(client: ExchangeClient, request: Any) -> None:
    # UpdateDraftRequest.attachments is None when the caller didn't touch attachments
    # at all, vs. a (possibly empty) list when they want to replace the whole set.
    if request.attachments is not None:
        _validate_attachments(request.attachments, client.settings)


def _validate_attachment_destination(client: ExchangeClient, request: Any) -> None:
    if request.save_path is not None:
        _validate_within_root([request.save_path], client.settings.attachment_root, "save_path")


list_emails = tool_handler("list_emails", ListEmailsRequest)
get_email = tool_handler("get_email", GetEmailRequest)
get_thread = tool_handler("get_thread", GetThreadRequest)
search_emails = tool_handler("search_emails", SearchEmailsRequest)
send_email = tool_handler("send_email", SendEmailRequest, before=_validate_outgoing_attachments)
reply_email = tool_handler("reply_email", ReplyEmailRequest, before=_validate_outgoing_attachments)
forward_email = tool_handler(
    "forward_email", ForwardEmailRequest, before=_validate_outgoing_attachments
)
move_email = tool_handler("move_email", FolderActionRequest)
copy_email = tool_handler("copy_email", FolderActionRequest)
delete_email = tool_handler("delete_email", DeleteEmailRequest)
mark_email = tool_handler("mark_email", MarkEmailRequest)
categorize_email = tool_handler("categorize_email", CategorizeEmailRequest)
list_categories = tool_handler("list_categories", ListCategoriesRequest)
list_folders = tool_handler("list_folders", ListFoldersRequest)
create_folder = tool_handler("create_folder", CreateFolderRequest)
rename_folder = tool_handler("rename_folder", RenameFolderRequest)
delete_folder = tool_handler("delete_folder", DeleteFolderRequest)
create_draft = tool_handler(
    "create_draft", DraftEmailRequest, before=_validate_outgoing_attachments
)
update_draft = tool_handler(
    "update_draft", UpdateDraftRequest, before=_validate_outgoing_attachments_if_set
)
send_draft = tool_handler("send_draft", SendDraftRequest)
get_attachment = tool_handler(
    "get_attachment", GetAttachmentRequest, before=_validate_attachment_destination
)


def _validate_within_root(paths: list[Path], root: Path | None, field: str) -> None:
    if not paths:
        return
    if root is None:
        raise APIError(
            "validation_error",
            f"{field} requires EXCHANGE_ATTACHMENT_ROOT to be configured",
            details=[
                {
                    "field": field,
                    "reason": "EXCHANGE_ATTACHMENT_ROOT is not set: local file access is "
                    f"refused by default: {path}",
                }
                for path in paths
            ],
        )
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
                    f"EXCHANGE_ATTACHMENT_MAX_COUNT={settings.attachment_max_count}",
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

    # Reject anything that isn't a plain file up front (devices, FIFOs, directories, ...).
    # This is an early, best-effort check for a friendly error message -- the authoritative
    # check happens against the opened file descriptor in _attach_files, since the file on
    # disk can still change between this stat and the actual read.
    not_regular = [str(path) for path in attachments if not path.is_file()]
    if not_regular:
        raise APIError(
            "validation_error",
            "one or more attachments are not regular files",
            details=[
                {"field": "attachments", "reason": f"not a regular file: {path}"}
                for path in not_regular
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
                    "reason": f"file exceeds EXCHANGE_ATTACHMENT_MAX_SIZE_MB="
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
                    f"EXCHANGE_ATTACHMENT_MAX_TOTAL_SIZE_MB={settings.attachment_max_total_size_mb}",
                }
            ],
        )
