from __future__ import annotations

from datetime import UTC, datetime
from types import SimpleNamespace

import pytest
from exchangelib.errors import ErrorAccessDenied, UnauthorizedError, UnknownTimeZone
from exchangelib.ewsdatetime import EWSTimeZone

from conftest import FakeExchangeBackend
from outlook_mcp.auth import build_auth_context
from outlook_mcp.config import Settings, get_settings
from outlook_mcp.errors import APIError, NotFoundError
from outlook_mcp.exchange_client import EWSExchangeBackend
from outlook_mcp.exchange_client.backend import build_default_backend
from outlook_mcp.models import (
    AddAttachmentRequest,
    BulkMoveEmailsRequest,
    CategorizeEmailRequest,
    CreateEventRequest,
    CreateFolderRequest,
    CreateInboxRuleActions,
    CreateInboxRuleConditions,
    CreateInboxRuleRequest,
    DeleteEmailRequest,
    DraftEmailRequest,
    FindFreeSlotsRequest,
    FolderActionRequest,
    ForwardEmailRequest,
    GetAttachmentRequest,
    ListEventsRequest,
    ReplyEmailRequest,
    SendDraftRequest,
    UpdateInboxRuleRequest,
    dump_model,
)


@pytest.fixture(autouse=True)
def _configure_attachment_root(settings, tmp_path) -> None:
    settings.attachment_root = tmp_path


_FOLDER = SimpleNamespace(id="folder-1", folder_class="IPF.Note")


def _item(**kwargs):
    kwargs.setdefault("id", "item-1")
    kwargs.setdefault("subject", "S")
    kwargs.setdefault("parent_folder_id", SimpleNamespace(id="folder-1"))
    return SimpleNamespace(**kwargs)


def _backend(settings, **fields) -> EWSExchangeBackend:
    fields.setdefault("default_timezone", EWSTimeZone("UTC"))
    backend = EWSExchangeBackend(settings)
    backend._account = SimpleNamespace(**fields)
    return backend


# --- модуль config/auth/backend ----------------------------------------------


def test_get_settings_reads_the_environment(monkeypatch) -> None:
    monkeypatch.setenv("EXCHANGE_SERVER", "https://mail.example.com/EWS/Exchange.asmx")
    monkeypatch.setenv("EXCHANGE_USERNAME", "user@example.com")
    monkeypatch.setenv("EXCHANGE_PASSWORD", "secret")
    get_settings.cache_clear()

    settings = get_settings()

    assert settings.exchange_username == "user@example.com"
    get_settings.cache_clear()


def test_auth_context_requires_a_derivable_smtp_address() -> None:
    # _env_file=None: the repo may hold a real .env whose EXCHANGE_EMAIL_ADDRESS
    # would satisfy the requirement this test is about.
    settings = Settings(
        _env_file=None,
        EXCHANGE_SERVER="https://mail.example.com/EWS/Exchange.asmx",
        EXCHANGE_USERNAME="DOMAIN\\user",
        EXCHANGE_PASSWORD="secret",
    )

    with pytest.raises(APIError) as excinfo:
        build_auth_context(settings)

    assert excinfo.value.details[0]["field"] == "EXCHANGE_EMAIL_ADDRESS"


def test_build_default_backend_returns_the_ews_backend(settings) -> None:
    assert isinstance(build_default_backend(settings), EWSExchangeBackend)


# --- base.py ------------------------------------------------------------------


def test_build_account_maps_a_connection_failure(settings, monkeypatch) -> None:
    import outlook_mcp.exchange_client.base as base_module

    class ExplodingAccount:
        def __init__(self, **kwargs):
            raise UnauthorizedError("bad creds")

    monkeypatch.setattr(base_module, "Account", ExplodingAccount)
    monkeypatch.setattr(settings, "exchange_email_address", "user@example.com")
    backend = EWSExchangeBackend(settings)

    with pytest.raises(APIError) as excinfo:
        _ = backend.account

    assert excinfo.value.code == "auth_failed"


def test_timezone_fallback_reraises_non_guid_ids(settings) -> None:
    backend = EWSExchangeBackend(settings)
    backend._configure_timezone_fallback()

    with pytest.raises(UnknownTimeZone):
        EWSTimeZone.from_ms_id("Definitely Not A Timezone")


def test_resolve_archive_falls_back_to_root_on_a_scoped_fault(settings, monkeypatch) -> None:
    import outlook_mcp.exchange_client.base as base_module
    from exchangelib.errors import ResponseMessageError

    class EmptyIdLookup:
        def __init__(self, account, folders):
            pass

        def resolve(self):
            return iter(())

    monkeypatch.setattr(base_module, "FolderCollection", EmptyIdLookup)
    root = SimpleNamespace(children=[])

    class Account:
        default_timezone = EWSTimeZone("UTC")
        inbox = _FOLDER
        sent = _FOLDER
        drafts = _FOLDER
        trash = _FOLDER
        junk = _FOLDER
        calendar = _FOLDER
        contacts = _FOLDER

        def __init__(self):
            self.root = root

        @property
        def archive_root(self):
            raise ResponseMessageError("no archive mailbox")

    backend = EWSExchangeBackend(settings)
    backend._account = Account()

    assert backend._resolve_folder("archive") is root


def test_map_exception_not_found_without_an_item_id(settings) -> None:
    from exchangelib.errors import ErrorItemNotFound

    backend = _backend(settings)

    mapped = backend._map_exception(ErrorItemNotFound("gone"))

    assert mapped.code == "not_found"
    assert "id" not in (mapped.extra or {})


def test_map_exception_transport_error_branch(settings) -> None:
    from exchangelib.errors import TransportError

    backend = _backend(settings)

    mapped = backend._map_exception(TransportError("socket closed"))

    assert mapped.code == "exchange_unavailable"


# --- calendar.py --------------------------------------------------------------


def test_update_event_dedupes_required_attendees_save_field(settings) -> None:
    item = _item(
        start=datetime(2026, 4, 8, 9, 0, tzinfo=UTC),
        end=datetime(2026, 4, 8, 10, 0, tzinfo=UTC),
        required_attendees=[
            SimpleNamespace(
                mailbox=SimpleNamespace(email_address="old@example.com"), response_type=None
            )
        ],
    )
    saves: list[dict] = []
    item.save = lambda **kwargs: saves.append(kwargs)
    backend = _backend(
        settings, calendar=SimpleNamespace(id="folder-1"), fetch=lambda **kwargs: iter([item])
    )

    from outlook_mcp.models import UpdateEventRequest

    backend.update_event(
        UpdateEventRequest(
            id="event-1",
            add_attendees=["new@example.com"],
            remove_attendees=["old@example.com"],
        )
    )

    assert saves[0]["update_fields"].count("required_attendees") == 1


def test_find_free_slots_maps_a_service_failure(settings) -> None:
    def get_free_busy_info(**kwargs):
        raise UnauthorizedError("bad creds")

    backend = _backend(settings, protocol=SimpleNamespace(get_free_busy_info=get_free_busy_info))

    with pytest.raises(APIError):
        backend.find_free_slots(
            FindFreeSlotsRequest(
                attendees=["user@example.com"],
                duration=60,
                start="2026-04-13T09:00:00+00:00",
                end="2026-04-13T10:00:00+00:00",
            )
        )


# --- contacts.py --------------------------------------------------------------


def test_gal_search_prepends_the_resolved_smtp_to_a_full_contact(settings, monkeypatch) -> None:
    import outlook_mcp.exchange_client.contacts as contacts_module

    contact = SimpleNamespace(
        id="contact-1",
        display_name="Ivan",
        file_as=None,
        email_addresses=[],
        phone_numbers=[],
    )
    mailbox = SimpleNamespace(email_address="smtp:ivan@corp.example.com", name="Ivan")

    class FakeResolveNames:
        def __init__(self, protocol):
            pass

        def call(self, **kwargs):
            yield (mailbox, contact)

    monkeypatch.setattr(contacts_module, "ResolveNames", FakeResolveNames)
    backend = _backend(settings, protocol=object())

    from outlook_mcp.models import SearchContactsRequest

    result = backend.search_contacts(SearchContactsRequest(query="ivan", source="gal"))

    assert result[0].email_addresses == ["ivan@corp.example.com"]


def test_gal_lookup_maps_a_resolve_failure(settings, monkeypatch) -> None:
    import outlook_mcp.exchange_client.contacts as contacts_module

    class FakeResolveNames:
        def __init__(self, protocol):
            pass

        def call(self, **kwargs):
            raise UnauthorizedError("bad creds")

    monkeypatch.setattr(contacts_module, "ResolveNames", FakeResolveNames)
    backend = _backend(settings, protocol=object())

    from outlook_mcp.models import GetContactRequest

    with pytest.raises(APIError):
        backend.get_contact(GetContactRequest(id="someone@example.com"))


# --- mailbox.py ---------------------------------------------------------------


def test_create_inbox_rule_resolves_the_move_target(settings, monkeypatch) -> None:
    import outlook_mcp.exchange_client.mailbox as mailbox_module

    captured: dict = {}

    class FakeCreate:
        def __init__(self, account):
            pass

        def call(self, rule, remove_outlook_rule_blob=True):
            captured["rule"] = rule
            yield from ()

    class FakeGet:
        def __init__(self, account):
            pass

        def call(self):
            yield from ()

    monkeypatch.setattr(mailbox_module, "CreateInboxRule", FakeCreate)
    monkeypatch.setattr(mailbox_module, "GetInboxRules", FakeGet)
    backend = _backend(settings)
    backend._resolve_folder = lambda value: SimpleNamespace(id="target-1", changekey="ck")

    backend.create_inbox_rule(
        CreateInboxRuleRequest(
            display_name="Move it",
            conditions=CreateInboxRuleConditions(has_attachments=True),
            actions=CreateInboxRuleActions(move_to_folder="Входящие/Проекты"),
        )
    )

    assert captured["rule"].actions.move_to_folder.folder_id.id == "target-1"


def test_update_inbox_rule_skips_error_and_none_rule_elements(settings, monkeypatch) -> None:
    import outlook_mcp.exchange_client.mailbox as mailbox_module

    target = SimpleNamespace(
        id="rule-1",
        display_name="R",
        priority=1,
        is_enabled=True,
        is_not_supported=False,
        conditions=None,
        actions=None,
    )

    class FakeGet:
        def __init__(self, account):
            pass

        def call(self):
            yield None
            yield ErrorAccessDenied("broken rule slot")
            yield target

    class FakeSet:
        def __init__(self, account):
            pass

        def call(self, rule, remove_outlook_rule_blob=True):
            yield from ()

    monkeypatch.setattr(mailbox_module, "GetInboxRules", FakeGet)
    monkeypatch.setattr(mailbox_module, "SetInboxRule", FakeSet)
    backend = _backend(settings)

    result = backend.update_inbox_rule(UpdateInboxRuleRequest(id="rule-1", is_enabled=False))

    assert result.id == "rule-1"


# --- email.py -----------------------------------------------------------------


def _mail_item(**kwargs):
    kwargs.setdefault("author", SimpleNamespace(email_address="a@example.com", name="A"))
    kwargs.setdefault("to_recipients", [])
    kwargs.setdefault("datetime_received", datetime(2026, 4, 7, 10, 0, tzinfo=UTC))
    kwargs.setdefault("is_read", True)
    kwargs.setdefault("categories", [])
    kwargs.setdefault("attachments", [])
    return _item(**kwargs)


def test_attachment_metadata_maps_every_field(settings) -> None:
    backend = _backend(settings)
    attachment = SimpleNamespace(
        attachment_id=SimpleNamespace(id="att-1"),
        name="report.pdf",
        size=123,
        content_type="application/pdf",
    )

    meta = backend._attachment_metadata(attachment)

    assert (meta.id, meta.name, meta.size) == ("att-1", "report.pdf", 123)
    assert meta.downloadable is True


def test_attach_files_enforces_the_total_size_budget(settings, tmp_path, monkeypatch) -> None:
    monkeypatch.setattr(settings, "attachment_max_total_size_mb", 1)
    one_mb = b"x" * (1024 * 1024)
    first = tmp_path / "a.bin"
    second = tmp_path / "b.bin"
    first.write_bytes(one_mb)
    second.write_bytes(one_mb)
    backend = _backend(settings)
    message = SimpleNamespace(attach=lambda a: None)

    with pytest.raises(APIError) as excinfo:
        backend._attach_files(message, [first, second])

    assert excinfo.value.code == "validation_error"


def test_sanitize_attachment_filename_fallbacks(settings) -> None:
    backend = _backend(settings)
    assert backend._sanitize_attachment_filename("..") == "attachment.bin"
    assert backend._sanitize_attachment_filename(None) == "attachment.bin"


def test_thread_items_returns_conversation_matches(settings) -> None:
    backend = _backend(settings)
    match = _mail_item(subject="Hello", conversation_id=SimpleNamespace(id="conv-1"))

    class Folder:
        def filter(self, *args, **kwargs):
            return self

        def only(self, *fields):
            return self

        def order_by(self, *fields):
            return self

        def __getitem__(self, item):
            return [match]

    items = backend._thread_items(Folder(), "conv-1", None, 20)

    assert items == [match]


def test_thread_items_without_conversation_or_subject_is_empty(settings) -> None:
    backend = _backend(settings)

    assert backend._thread_items(object(), None, "   ", 20) == []


def test_send_email_reports_the_saved_copy_id(settings, monkeypatch) -> None:
    import outlook_mcp.exchange_client.email as email_module

    class SendingMessage(SimpleNamespace):
        def __init__(self, **kwargs):
            kwargs.pop("account", None)
            kwargs.pop("folder", None)
            super().__init__(**kwargs)
            self.id = "sent-1"

        def send_and_save(self):
            pass

    monkeypatch.setattr(email_module, "Message", SendingMessage)
    backend = _backend(settings, drafts=_FOLDER, sent=_FOLDER)

    from outlook_mcp.models import SendEmailRequest

    result = backend.send_email(SendEmailRequest(to=["u@example.com"], subject="S", body="B"))

    assert result.id == "sent-1"


def test_reply_and_forward_map_creation_failures(settings) -> None:
    item = _mail_item()
    item.create_reply = lambda **kwargs: (_ for _ in ()).throw(ErrorAccessDenied("no reply"))
    item.create_forward = lambda **kwargs: (_ for _ in ()).throw(ErrorAccessDenied("no fwd"))
    backend = _backend(settings, fetch=lambda **kwargs: iter([item]))

    with pytest.raises(APIError):
        backend.reply_email(ReplyEmailRequest(id="item-1", body="B"))
    with pytest.raises(APIError):
        backend.forward_email(ForwardEmailRequest(id="item-1", to=["u@example.com"]))


def test_move_copy_delete_map_failures(settings, monkeypatch) -> None:
    item = _mail_item()
    item.move = lambda to_folder: (_ for _ in ()).throw(ErrorAccessDenied("locked"))
    item.copy = lambda to_folder: (_ for _ in ()).throw(ErrorAccessDenied("locked"))
    item.move_to_trash = lambda: (_ for _ in ()).throw(ErrorAccessDenied("locked"))
    backend = _backend(settings, fetch=lambda **kwargs: iter([item]))
    backend._resolve_folder = lambda value: _FOLDER

    with pytest.raises(APIError):
        backend.move_email(FolderActionRequest(id="item-1", folder="inbox"))
    with pytest.raises(APIError):
        backend.copy_email(FolderActionRequest(id="item-1", folder="inbox"))
    with pytest.raises(APIError):
        backend.delete_email(DeleteEmailRequest(id="item-1"))


def test_bulk_move_maps_a_wholesale_failure(settings) -> None:
    backend = _backend(settings)
    backend._resolve_folder = lambda value: _FOLDER
    backend._account.bulk_move = lambda ids, to_folder: (_ for _ in ()).throw(
        UnauthorizedError("bad creds")
    )

    with pytest.raises(APIError):
        backend.move_emails(BulkMoveEmailsRequest(ids=["a"], folder="inbox"))


def test_categorize_email_maps_a_save_failure(settings) -> None:
    item = _mail_item(categories=["Old"])
    item.save = lambda **kwargs: (_ for _ in ()).throw(ErrorAccessDenied("locked"))
    backend = _backend(settings, fetch=lambda **kwargs: iter([item]))

    with pytest.raises(APIError):
        backend.categorize_email(CategorizeEmailRequest(id="item-1", categories=["New"]))


def test_create_folder_reports_the_new_path(settings, monkeypatch) -> None:
    import outlook_mcp.exchange_client.email as email_module

    class SavedFolder(SimpleNamespace):
        def __init__(self, **kwargs):
            kwargs.pop("parent", None)
            super().__init__(**kwargs)
            self.id = "new-folder"
            self.parent = None

        def save(self):
            pass

    monkeypatch.setattr(email_module, "Folder", SavedFolder)
    backend = _backend(settings)
    backend._resolve_folder = lambda value: _FOLDER

    result = backend.create_folder(CreateFolderRequest(name="Projects", parent="inbox"))

    assert result.status == "created"
    assert result.id == "new-folder"


def test_create_draft_and_send_draft_map_failures(settings, monkeypatch) -> None:
    import outlook_mcp.exchange_client.email as email_module

    class FailingMessage(SimpleNamespace):
        def __init__(self, **kwargs):
            kwargs.pop("account", None)
            kwargs.pop("folder", None)
            super().__init__(**kwargs)

        def save(self):
            raise ErrorAccessDenied("no drafts")

    monkeypatch.setattr(email_module, "Message", FailingMessage)
    backend = _backend(settings, drafts=_FOLDER)

    with pytest.raises(APIError):
        backend.create_draft(DraftEmailRequest(to=["u@example.com"], subject="S", body="B"))

    draft = _mail_item()
    draft.send = lambda **kwargs: (_ for _ in ()).throw(ErrorAccessDenied("cannot send"))
    backend._account.fetch = lambda **kwargs: iter([draft])
    with pytest.raises(APIError):
        backend.send_draft(SendDraftRequest(id="item-1"))


def test_add_attachment_absolute_path_reraises_api_errors(settings, tmp_path) -> None:
    source = tmp_path / "note.txt"
    source.write_text("hi")
    item = _mail_item()

    def attach(a):
        raise APIError("validation_error", "custom validation failure")

    item.attach = attach
    backend = _backend(settings, fetch=lambda **kwargs: iter([item]))

    with pytest.raises(APIError) as excinfo:
        backend.add_attachment(AddAttachmentRequest(email_id="item-1", path=source))

    assert excinfo.value.message == "custom validation failure"


def test_get_attachment_rejects_item_attachments(settings, tmp_path, monkeypatch) -> None:
    from exchangelib import ItemAttachment

    monkeypatch.setattr(settings, "attachment_root", tmp_path)
    attachment = ItemAttachment()
    attachment.attachment_id = SimpleNamespace(id="att-1")
    attachment.name = "embedded mail"
    item = _mail_item(attachments=[attachment])
    backend = _backend(settings, fetch=lambda **kwargs: iter([item]))

    with pytest.raises(APIError) as excinfo:
        backend.get_attachment(
            GetAttachmentRequest(email_id="item-1", attachment_id="att-1", save_path=tmp_path)
        )

    assert excinfo.value.code == "validation_error"


def test_get_attachment_cleans_up_after_a_failed_write(settings, tmp_path, monkeypatch) -> None:
    from exchangelib import FileAttachment

    monkeypatch.setattr(settings, "attachment_root", tmp_path)

    class BrokenContent(FileAttachment):
        @property
        def content(self):
            raise ErrorAccessDenied("content is gone")

    attachment = BrokenContent()
    attachment.attachment_id = SimpleNamespace(id="att-1")
    attachment.name = "report.txt"
    item = _mail_item(attachments=[attachment])
    backend = _backend(settings, fetch=lambda **kwargs: iter([item]))

    with pytest.raises(APIError):
        backend.get_attachment(
            GetAttachmentRequest(email_id="item-1", attachment_id="att-1", save_path=tmp_path)
        )

    assert list(tmp_path.iterdir()) == []


def test_get_attachment_unknown_id_is_not_found(settings, tmp_path, monkeypatch) -> None:
    monkeypatch.setattr(settings, "attachment_root", tmp_path)
    item = _mail_item(attachments=[])
    backend = _backend(settings, fetch=lambda **kwargs: iter([item]))

    with pytest.raises(NotFoundError):
        backend.get_attachment(
            GetAttachmentRequest(email_id="item-1", attachment_id="nope", save_path=tmp_path)
        )


# --- models.py / server.py / tools --------------------------------------------


def test_server_address_passes_non_strings_through() -> None:
    from outlook_mcp.models import _server_address

    marker = object()
    assert _server_address(marker) is marker


def test_categorize_request_rejects_blank_names() -> None:
    with pytest.raises(ValueError, match="must not be blank"):
        CategorizeEmailRequest(id="email-1", categories=["  "])


def test_range_validators_reject_reversed_windows() -> None:
    with pytest.raises(ValueError):
        ListEventsRequest(start="2026-04-13T10:00:00+00:00", end="2026-04-13T09:00:00+00:00")
    with pytest.raises(ValueError):
        CreateEventRequest(
            subject="S", start="2026-04-13T10:00:00+00:00", end="2026-04-13T09:00:00+00:00"
        )
    with pytest.raises(ValueError):
        FindFreeSlotsRequest(
            attendees=["u@example.com"],
            duration=60,
            start="2026-04-13T10:00:00+00:00",
            end="2026-04-13T09:00:00+00:00",
        )


def test_dump_model_passes_dicts_and_scalars_through() -> None:
    assert dump_model({"a": 1}) == {"a": 1}
    assert dump_model("scalar") == "scalar"


def test_configure_logging_writes_to_a_file(tmp_path, settings, monkeypatch) -> None:
    import logging

    from outlook_mcp.server import configure_logging

    log_file = tmp_path / "server.log"
    monkeypatch.setattr(settings, "log_file", log_file)

    configure_logging(settings)
    logging.getLogger("outlook_mcp.test").warning("hello file")
    for handler in logging.getLogger().handlers:
        handler.flush()

    assert log_file.exists()


def test_main_refuses_an_unknown_transport(monkeypatch, settings) -> None:
    import outlook_mcp.server as server_module

    monkeypatch.setattr(settings, "mcp_transport", "carrier-pigeon")
    monkeypatch.setattr(server_module, "get_settings", lambda: settings)
    monkeypatch.setattr(server_module, "build_mcp_server", lambda settings=None: SimpleNamespace())

    with pytest.raises(RuntimeError, match="unsupported transport"):
        server_module.main()


def test_added_attachment_hook_resolves_relative_paths(tmp_path, settings, monkeypatch) -> None:
    from outlook_mcp.exchange_client import ExchangeClient
    from outlook_mcp.tools.email import _validate_added_attachment

    monkeypatch.setattr(settings, "attachment_root", tmp_path)
    (tmp_path / "note.txt").write_text("hi")
    client = ExchangeClient(settings=settings, backend=FakeExchangeBackend())

    # Does not raise: the relative path resolves under the root.
    _validate_added_attachment(client, AddAttachmentRequest(email_id="email-1", path="note.txt"))


def test_update_event_request_and_free_slot_reject_reversed_ranges() -> None:
    from outlook_mcp.models import FreeSlot, UpdateEventRequest

    with pytest.raises(ValueError):
        UpdateEventRequest(
            id="event-1", start="2026-04-13T10:00:00+00:00", end="2026-04-13T09:00:00+00:00"
        )
    with pytest.raises(ValueError):
        FreeSlot(start="2026-04-13T10:00:00+00:00", end="2026-04-13T09:00:00+00:00")


def test_update_event_remove_attendees_alone_still_saves_the_field(settings) -> None:
    item = _item(
        start=datetime(2026, 4, 8, 9, 0, tzinfo=UTC),
        end=datetime(2026, 4, 8, 10, 0, tzinfo=UTC),
        required_attendees=[
            SimpleNamespace(
                mailbox=SimpleNamespace(email_address="old@example.com"), response_type=None
            )
        ],
    )
    saves: list[dict] = []
    item.save = lambda **kwargs: saves.append(kwargs)
    backend = _backend(
        settings, calendar=SimpleNamespace(id="folder-1"), fetch=lambda **kwargs: iter([item])
    )

    from outlook_mcp.models import UpdateEventRequest

    backend.update_event(UpdateEventRequest(id="event-1", remove_attendees=["old@example.com"]))

    assert saves[0]["update_fields"] == ["required_attendees"]
    assert item.required_attendees == []


def test_bulk_copy_and_bulk_delete_map_wholesale_failures(settings) -> None:
    from outlook_mcp.models import BulkDeleteEmailsRequest

    backend = _backend(settings)
    backend._resolve_folder = lambda value: _FOLDER
    backend._account.bulk_copy = lambda ids, to_folder: (_ for _ in ()).throw(
        UnauthorizedError("bad creds")
    )
    backend._account.bulk_delete = lambda ids, delete_type: (_ for _ in ()).throw(
        UnauthorizedError("bad creds")
    )

    with pytest.raises(APIError):
        backend.copy_emails(BulkMoveEmailsRequest(ids=["a"], folder="inbox"))
    with pytest.raises(APIError):
        backend.delete_emails(BulkDeleteEmailsRequest(ids=["a"]))


def test_add_attachment_resolves_relative_paths_against_the_root(
    settings, tmp_path, monkeypatch
) -> None:
    monkeypatch.setattr(settings, "attachment_root", tmp_path)
    (tmp_path / "note.txt").write_text("hi")
    attached: list = []
    item = _mail_item()
    item.attach = lambda a: attached.append(a)
    backend = _backend(settings, fetch=lambda **kwargs: iter([item]))

    result = backend.add_attachment(AddAttachmentRequest(email_id="item-1", path="note.txt"))

    assert result.status == "updated"
    assert len(attached) == 1


def test_unique_path_reraises_unexpected_open_errors(settings, tmp_path, monkeypatch) -> None:
    import errno as errno_module
    import os

    real_open = os.open

    def denied(path, flags, *args, **kwargs):
        if str(path).endswith("denied.txt"):
            raise OSError(errno_module.EACCES, "permission denied")
        return real_open(path, flags, *args, **kwargs)

    monkeypatch.setattr(os, "open", denied)
    backend = _backend(settings)

    with pytest.raises(OSError):
        backend._create_new_file(tmp_path / "denied.txt")


def test_timezone_fallback_double_checked_lock_inner_branch(settings, monkeypatch) -> None:
    """The race the lock exists for: another thread completes the patch between
    the outer check and acquiring the lock. Simulated with a lock whose acquire
    flips the flag, so the inner check must return without re-patching."""
    import outlook_mcp.exchange_client.base as base_module

    class FlagFlippingLock:
        def __enter__(self):
            base_module._TIMEZONE_FALLBACK_PATCHED = True

        def __exit__(self, *exc):
            return False

    monkeypatch.setattr(base_module, "_TIMEZONE_FALLBACK_PATCHED", False)
    monkeypatch.setattr(base_module, "_TIMEZONE_FALLBACK_LOCK", FlagFlippingLock())
    original = EWSTimeZone.from_ms_id

    EWSExchangeBackend(settings)._configure_timezone_fallback()

    # The inner check won: from_ms_id was not wrapped a second time.
    assert EWSTimeZone.from_ms_id == original
