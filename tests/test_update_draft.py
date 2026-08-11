from __future__ import annotations

from types import SimpleNamespace

import pytest

from outlook_mcp.exchange_client import EWSExchangeBackend
from outlook_mcp.models import ActionResult, UpdateDraftRequest


def _fake_draft(events: list[tuple]) -> SimpleNamespace:
    class FakeDraft:
        id = "draft-1"
        subject = "Old subject"
        body = "Old body"
        to_recipients = []
        cc_recipients = []
        bcc_recipients = []
        attachments = []
        parent_folder_id = SimpleNamespace(id="drafts-folder")

        def detach(self, attachments):
            events.append(("detach", list(attachments)))
            self.attachments = [a for a in self.attachments if a not in attachments]

        def attach(self, attachment):
            events.append(("attach", attachment.name))
            self.attachments = [*self.attachments, attachment]

        def save(self, update_fields=None):
            events.append(("save", update_fields))

    return FakeDraft()


def _backend_with_draft(draft: SimpleNamespace) -> EWSExchangeBackend:
    backend = EWSExchangeBackend.__new__(EWSExchangeBackend)
    backend.settings = SimpleNamespace(
        attachment_root=None,
        attachment_max_count=10,
        attachment_max_size_mb=25,
        attachment_max_total_size_mb=25,
        exchange_timezone_fallback=None,
    )

    def fetch(ids, folder=None):
        yield draft

    drafts_folder = SimpleNamespace(id="drafts-folder")
    backend._account = SimpleNamespace(fetch=fetch, drafts=drafts_folder)
    return backend


def test_update_draft_updates_only_provided_fields() -> None:
    events: list[tuple] = []
    draft = _fake_draft(events)
    backend = _backend_with_draft(draft)

    result = backend.update_draft(UpdateDraftRequest(id="draft-1", subject="New subject"))

    assert draft.subject == "New subject"
    assert draft.body == "Old body"
    assert events == [("save", ["subject"])]
    assert result == ActionResult(id="draft-1", status="updated", updated_fields=["subject"])


def test_update_draft_updates_recipients_and_body() -> None:
    events: list[tuple] = []
    draft = _fake_draft(events)
    backend = _backend_with_draft(draft)

    result = backend.update_draft(
        UpdateDraftRequest(
            id="draft-1",
            to=["a@example.com", "b@example.com"],
            cc=[],
            body="New body",
        )
    )

    assert [m.email_address for m in draft.to_recipients] == ["a@example.com", "b@example.com"]
    assert draft.cc_recipients == []
    assert draft.body == "New body"
    assert events == [("save", ["to_recipients", "body", "cc_recipients"])]
    assert set(result.updated_fields) == {"to", "cc", "body"}


def test_update_draft_with_no_fields_does_not_save() -> None:
    events: list[tuple] = []
    draft = _fake_draft(events)
    backend = _backend_with_draft(draft)

    result = backend.update_draft(UpdateDraftRequest(id="draft-1"))

    assert events == []
    assert result == ActionResult(id="draft-1", status="updated", updated_fields=[])


def test_update_draft_replaces_attachments(tmp_path) -> None:
    events: list[tuple] = []
    draft = _fake_draft(events)
    old_attachment = object()
    draft.attachments = [old_attachment]
    backend = _backend_with_draft(draft)

    new_file = tmp_path / "note.txt"
    new_file.write_text("hi")

    result = backend.update_draft(UpdateDraftRequest(id="draft-1", attachments=[new_file]))

    assert events[0] == ("detach", [old_attachment])
    assert [a.name for a in draft.attachments] == ["note.txt"]
    assert result.updated_fields == ["attachments"]


def test_update_draft_clears_attachments_with_empty_list() -> None:
    events: list[tuple] = []
    draft = _fake_draft(events)
    old_attachment = object()
    draft.attachments = [old_attachment]
    backend = _backend_with_draft(draft)

    result = backend.update_draft(UpdateDraftRequest(id="draft-1", attachments=[]))

    assert events == [("detach", [old_attachment])]
    assert draft.attachments == []
    assert result.updated_fields == ["attachments"]


def test_update_draft_html_body_wraps_in_html_body() -> None:
    events: list[tuple] = []
    draft = _fake_draft(events)
    backend = _backend_with_draft(draft)

    backend.update_draft(UpdateDraftRequest(id="draft-1", body="<p>Hi</p>", body_type="html"))

    assert str(draft.body) == "<p>Hi</p>"
    from exchangelib import HTMLBody

    assert isinstance(draft.body, HTMLBody)


def test_update_draft_rejects_wrong_folder() -> None:
    events: list[tuple] = []
    draft = _fake_draft(events)
    draft.parent_folder_id = SimpleNamespace(id="some-other-folder")
    backend = _backend_with_draft(draft)

    with pytest.raises(Exception):
        backend.update_draft(UpdateDraftRequest(id="draft-1", subject="New subject"))


def test_update_draft_maps_exceptions_from_save() -> None:
    events: list[tuple] = []
    draft = _fake_draft(events)

    def failing_save(update_fields=None):
        raise RuntimeError("boom")

    draft.save = failing_save
    backend = _backend_with_draft(draft)

    with pytest.raises(Exception):
        backend.update_draft(UpdateDraftRequest(id="draft-1", subject="New subject"))


def test_validate_outgoing_attachments_if_set_skips_when_attachments_omitted() -> None:
    from types import SimpleNamespace as NS

    from outlook_mcp.tools.email import _validate_outgoing_attachments_if_set

    client = NS(settings=NS(attachment_root=None))
    # attachments is None (field omitted) -- must not raise even without
    # EXCHANGE_ATTACHMENT_ROOT configured.
    _validate_outgoing_attachments_if_set(client, UpdateDraftRequest(id="draft-1"))


def test_validate_outgoing_attachments_if_set_validates_when_attachments_given() -> None:
    from types import SimpleNamespace as NS

    from outlook_mcp.errors import APIError
    from outlook_mcp.tools.email import _validate_outgoing_attachments_if_set

    client = NS(settings=NS(attachment_root=None))
    request = UpdateDraftRequest(id="draft-1", attachments=["note.txt"])

    with pytest.raises(APIError):
        _validate_outgoing_attachments_if_set(client, request)
