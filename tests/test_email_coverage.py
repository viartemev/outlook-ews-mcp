from __future__ import annotations

from datetime import UTC, datetime
from types import SimpleNamespace

import pytest
from exchangelib.errors import ErrorAccessDenied, UnauthorizedError

from outlook_mcp.errors import APIError, NotFoundError
from outlook_mcp.exchange_client import EWSExchangeBackend
from outlook_mcp.models import (
    AddAttachmentRequest,
    BulkDeleteEmailsRequest,
    BulkMoveEmailsRequest,
    CreateFolderRequest,
    DeleteAttachmentRequest,
    DeleteEmailRequest,
    DraftEmailRequest,
    FolderActionRequest,
    GetEmailRequest,
    ListCategoriesRequest,
    ListEmailsRequest,
    ListFoldersRequest,
    MarkEmailRequest,
    SendDraftRequest,
    SendEmailRequest,
)

_FOLDER_ID = "folder-1"


class FakeQuerySet:
    def __init__(self, items, explode: Exception | None = None):
        self.items = list(items)
        self.explode = explode
        self.filters: list = []

    def all(self):
        return self

    def filter(self, *args, **kwargs):
        self.filters.append(kwargs or args)
        return self

    def only(self, *fields):
        return self

    def order_by(self, *fields):
        return self

    def __getitem__(self, item):
        if self.explode is not None:
            raise self.explode
        return self.items


class FakeFolder(FakeQuerySet):
    def __init__(self, items=(), name="Входящие", children=(), explode=None):
        super().__init__(items, explode)
        self.id = _FOLDER_ID
        self.name = name
        self.children = list(children)
        self.parent = None
        self.unread_count = 0
        self.total_count = len(self.items)


class FakeItem(SimpleNamespace):
    def __init__(self, **kwargs):
        kwargs.setdefault("id", "email-1")
        kwargs.setdefault("subject", "Hello")
        kwargs.setdefault("parent_folder_id", SimpleNamespace(id=_FOLDER_ID))
        kwargs.setdefault("datetime_received", datetime(2026, 4, 7, 10, 0, tzinfo=UTC))
        kwargs.setdefault("is_read", True)
        kwargs.setdefault("categories", [])
        kwargs.setdefault("author", SimpleNamespace(email_address="a@example.com", name="A"))
        kwargs.setdefault("to_recipients", [])
        kwargs.setdefault("attachments", [])
        super().__init__(**kwargs)
        self.saved: list = []
        self.deleted: list = []
        self.trashed = False

    def save(self, **kwargs):
        self.saved.append(kwargs)

    def delete(self, **kwargs):
        self.deleted.append(kwargs)

    def move_to_trash(self):
        self.trashed = True


def _backend(settings, *, folder=None, item=None, **account_extra) -> EWSExchangeBackend:
    from exchangelib.ewsdatetime import EWSTimeZone

    backend = EWSExchangeBackend(settings)
    inbox = folder if folder is not None else FakeFolder()
    fields = {
        "default_timezone": EWSTimeZone("UTC"),
        "inbox": inbox,
        "drafts": FakeFolder(name="Черновики"),
        # _resolve_folder builds its builtin map eagerly, touching all of these.
        "root": SimpleNamespace(children=[], tois=SimpleNamespace(children=[inbox])),
        "sent": FakeFolder(name="Отправленные"),
        "trash": FakeFolder(name="Удаленные"),
        "junk": FakeFolder(name="Нежелательная"),
        "calendar": FakeFolder(name="Календарь"),
        "contacts": FakeFolder(name="Контакты"),
    }
    fields.update(account_extra)
    if item is not None:
        fields.setdefault("fetch", lambda **kwargs: iter([item]))
    backend._account = SimpleNamespace(**fields)
    return backend


def test_list_emails_maps_every_filter(settings) -> None:
    folder = FakeFolder([FakeItem()])
    backend = _backend(settings, folder=folder)

    backend.list_emails(
        ListEmailsRequest(
            from_address="a@example.com",
            subject="report",
            since="2026-04-01",
            before="2026-04-30",
            unread_only=True,
            has_attachments=True,
        )
    )

    merged: dict = {}
    for f in folder.filters:
        if isinstance(f, dict):
            merged.update(f)
    assert merged["author__iexact"] == "a@example.com"
    assert merged["subject__icontains"] == "report"
    assert merged["is_read"] is False
    assert merged["has_attachments"] is True
    assert merged["datetime_received__gte"].date().isoformat() == "2026-04-01"
    # "before" is inclusive: the cutoff is the start of the next day.
    assert merged["datetime_received__lt"].date().isoformat() == "2026-05-01"


def test_list_emails_maps_a_server_failure(settings) -> None:
    backend = _backend(settings, folder=FakeFolder(explode=UnauthorizedError("bad creds")))

    with pytest.raises(APIError) as excinfo:
        backend.list_emails(ListEmailsRequest())

    assert excinfo.value.code == "auth_failed"


def test_get_email_truncates_an_oversized_body(settings) -> None:
    big = "x" * (settings.email_body_max_chars + 100)
    item = FakeItem(text_body=big, body=f"<html>{big}</html>")
    backend = _backend(settings, item=item)

    result = backend.get_email(GetEmailRequest(id="email-1"))

    assert result.truncated is True
    assert len(result.body_text) == settings.email_body_max_chars


def test_send_email_maps_a_send_failure(settings, monkeypatch) -> None:
    import outlook_mcp.exchange_client.email as email_module

    class ExplodingMessage(FakeItem):
        def __init__(self, **kwargs):
            kwargs.pop("account", None)
            kwargs.pop("folder", None)
            super().__init__(**kwargs)

        def send_and_save(self):
            raise ErrorAccessDenied("relay denied")

    monkeypatch.setattr(email_module, "Message", ExplodingMessage)
    backend = _backend(settings)

    with pytest.raises(APIError) as excinfo:
        backend.send_email(SendEmailRequest(to=["u@example.com"], subject="Hi", body="B"))

    assert excinfo.value.code == "permission_denied"


def test_copy_email_returns_both_ids(settings) -> None:
    item = FakeItem()
    item.copy = lambda to_folder: ("copy-1", "ck")
    backend = _backend(settings, folder=FakeFolder(), item=item)

    result = backend.copy_email(FolderActionRequest(id="email-1", folder="inbox"))

    assert result.new_id == "copy-1"
    assert result.status == "copied"


def test_copy_email_tolerates_a_missing_new_id(settings) -> None:
    item = FakeItem()
    item.copy = lambda to_folder: None
    backend = _backend(settings, folder=FakeFolder(), item=item)

    assert backend.copy_email(FolderActionRequest(id="email-1", folder="inbox")).new_id is None


def test_delete_email_soft_uses_the_trash(settings) -> None:
    item = FakeItem()
    backend = _backend(settings, item=item)

    backend.delete_email(DeleteEmailRequest(id="email-1"))

    assert item.trashed is True
    assert item.deleted == []


def test_delete_email_hard_bypasses_the_trash(settings) -> None:
    item = FakeItem()
    backend = _backend(settings, item=item)

    backend.delete_email(DeleteEmailRequest(id="email-1", hard_delete=True))

    assert item.trashed is False
    assert item.deleted == [{}]


def test_mark_email_sets_importance(settings) -> None:
    item = FakeItem(importance="Normal")
    backend = _backend(settings, item=item)

    backend.mark_email(MarkEmailRequest(id="email-1", importance="high"))

    assert item.importance == "High"
    assert item.saved[-1]["update_fields"] == ["importance"]


def test_mark_email_maps_a_save_failure(settings) -> None:
    item = FakeItem(is_read=False)
    item.save = lambda **kwargs: (_ for _ in ()).throw(ErrorAccessDenied("locked"))
    backend = _backend(settings, item=item)

    with pytest.raises(APIError) as excinfo:
        backend.mark_email(MarkEmailRequest(id="email-1", read=True))

    assert excinfo.value.code == "permission_denied"


def _bulk_account(results):
    return {
        "bulk_move": lambda ids, to_folder: results,
        "bulk_copy": lambda ids, to_folder: results,
        "bulk_delete": lambda ids, delete_type: results,
    }


def test_bulk_copy_and_delete_partition_successes_and_failures(settings) -> None:
    results = [("new-1", "ck"), ErrorAccessDenied("locked")]
    backend = _backend(settings, folder=FakeFolder())
    backend._account.bulk_copy = lambda ids, to_folder: results
    backend._account.bulk_delete = lambda ids, delete_type: [None, ErrorAccessDenied("x")]

    copied = backend.copy_emails(BulkMoveEmailsRequest(ids=["a", "b"], folder="inbox"))
    assert [r.id for r in copied.succeeded] == ["a"]
    assert [f.id for f in copied.failed] == ["b"]

    deleted = backend.delete_emails(BulkDeleteEmailsRequest(ids=["a", "b"]))
    assert [r.id for r in deleted.succeeded] == ["a"]
    assert [f.id for f in deleted.failed] == ["b"]


def test_bulk_operations_map_a_wholesale_failure(settings) -> None:
    backend = _backend(settings, folder=FakeFolder())
    backend._account.bulk_move = lambda ids, to_folder: (_ for _ in ()).throw(
        UnauthorizedError("bad creds")
    )

    with pytest.raises(APIError) as excinfo:
        backend.move_emails(BulkMoveEmailsRequest(ids=["a"], folder="inbox"))

    assert excinfo.value.code == "auth_failed"


def test_thread_subject_fallback_maps_a_search_failure(settings) -> None:
    backend = _backend(settings)
    folder = FakeFolder(explode=UnauthorizedError("bad creds"))

    with pytest.raises(APIError) as excinfo:
        backend._thread_items(folder, None, "Re: Hello", 20)

    assert excinfo.value.code == "auth_failed"


def test_list_folders_reports_children_to_the_requested_depth(settings) -> None:
    leaf = FakeFolder(name="Leaf")
    child = FakeFolder(name="Child", children=[leaf])
    backend = _backend(settings, folder=FakeFolder(name="Входящие", children=[child]))

    result = backend.list_folders(ListFoldersRequest(parent="inbox", depth=2))

    assert result[0].name == "Child"
    assert result[0].children[0].name == "Leaf"


def test_create_folder_maps_a_save_failure(settings, monkeypatch) -> None:
    import outlook_mcp.exchange_client.email as email_module

    class ExplodingFolder(FakeItem):
        def __init__(self, **kwargs):
            kwargs.pop("parent", None)
            super().__init__(**kwargs)

        def save(self, **kwargs):
            raise ErrorAccessDenied("no rights here")

    monkeypatch.setattr(email_module, "Folder", ExplodingFolder)
    backend = _backend(settings, folder=FakeFolder())

    with pytest.raises(APIError) as excinfo:
        backend.create_folder(CreateFolderRequest(name="X", parent="inbox"))

    assert excinfo.value.code == "permission_denied"


def test_create_draft_and_send_draft(settings, monkeypatch) -> None:
    import outlook_mcp.exchange_client.email as email_module

    class DraftMessage(FakeItem):
        def __init__(self, **kwargs):
            kwargs.pop("account", None)
            kwargs.pop("folder", None)
            kwargs.setdefault("id", "draft-1")
            super().__init__(**kwargs)

        def send_and_save(self):
            self.sent = True

    monkeypatch.setattr(email_module, "Message", DraftMessage)
    backend = _backend(settings)

    created = backend.create_draft(DraftEmailRequest(to=["u@example.com"], subject="S", body="B"))
    assert created.status == "draft"

    draft = DraftMessage()
    backend._account.fetch = lambda **kwargs: iter([draft])
    sent = backend.send_draft(SendDraftRequest(id="draft-1"))
    # The draft's id dies the moment it is sent; it must not be echoed back.
    assert sent.id is None
    assert draft.sent is True


def test_add_attachment_requires_a_root_for_relative_paths(settings) -> None:
    backend = _backend(settings)
    assert settings.attachment_root is None

    with pytest.raises(APIError) as excinfo:
        backend.add_attachment(AddAttachmentRequest(email_id="email-1", path="note.txt"))

    assert excinfo.value.code == "validation_error"


def test_add_attachment_names_a_missing_file(settings, tmp_path) -> None:
    backend = _backend(settings)

    with pytest.raises(APIError) as excinfo:
        backend.add_attachment(
            AddAttachmentRequest(email_id="email-1", path=tmp_path / "absent.txt")
        )

    assert excinfo.value.code == "validation_error"
    assert "absent.txt" in str(excinfo.value.details)


def test_add_attachment_maps_an_attach_failure(settings, tmp_path) -> None:
    source = tmp_path / "note.txt"
    source.write_text("hi")
    item = FakeItem()
    item.attach = lambda a: (_ for _ in ()).throw(ErrorAccessDenied("read only"))
    backend = _backend(settings.model_copy(update={"attachment_root": tmp_path}), item=item)

    with pytest.raises(APIError) as excinfo:
        backend.add_attachment(AddAttachmentRequest(email_id="email-1", path=source))

    assert excinfo.value.code == "permission_denied"


def test_delete_attachment_maps_a_detach_failure(settings) -> None:
    attachment = SimpleNamespace(attachment_id=SimpleNamespace(id="att-1"), name="x")
    item = FakeItem(attachments=[attachment])
    item.detach = lambda a: (_ for _ in ()).throw(ErrorAccessDenied("locked"))
    backend = _backend(settings, item=item)

    with pytest.raises(APIError) as excinfo:
        backend.delete_attachment(
            DeleteAttachmentRequest(email_id="email-1", attachment_id="att-1")
        )

    assert excinfo.value.code == "permission_denied"


def test_delete_attachment_reports_an_unknown_id_as_not_found(settings) -> None:
    item = FakeItem(attachments=[])
    backend = _backend(settings, item=item)

    with pytest.raises(NotFoundError):
        backend.delete_attachment(DeleteAttachmentRequest(email_id="email-1", attachment_id="nope"))


def test_list_categories_fallback_skips_blank_names_and_maps_failures(settings) -> None:
    backend = _backend(settings)
    backend._master_category_list = lambda: None
    good = FakeFolder([FakeItem(categories=["  ", "Ops"])])
    backend._resolve_folder = lambda value: good

    result = backend.list_categories(ListCategoriesRequest(folders=["inbox"]))
    assert [(u.name, u.count) for u in result] == [("Ops", 1)]

    backend._resolve_folder = lambda value: FakeFolder(explode=UnauthorizedError("bad creds"))
    with pytest.raises(APIError):
        backend.list_categories(ListCategoriesRequest(folders=["inbox"]))


def test_master_category_list_handles_missing_and_bad_color_values(settings) -> None:
    xml = (
        b'<categories xmlns="CategoryList.xsd">'
        b'<category name="A" color="oops" usageCount="1"/>'
        b"</categories>"
    )
    backend = _backend(
        settings,
        calendar=SimpleNamespace(get_user_configuration=lambda name: SimpleNamespace(xml_data=xml)),
    )

    result = backend._master_category_list()

    assert result[0].color is None

    backend._account.calendar = SimpleNamespace(
        get_user_configuration=lambda name: SimpleNamespace(xml_data=None)
    )
    assert backend._master_category_list() is None
