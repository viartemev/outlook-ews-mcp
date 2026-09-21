from __future__ import annotations

from datetime import UTC, datetime
from types import SimpleNamespace

import pytest
from exchangelib import HTMLBody, Mailbox
from pydantic import ValidationError

from outlook_mcp.errors import APIError
from outlook_mcp.exchange_client import EWSExchangeBackend
from outlook_mcp.mentions import render_mentions
from outlook_mcp.models import (
    CreateReplyDraftRequest,
    DraftEmailRequest,
    Mention,
    ReplyEmailRequest,
    SendDraftRequest,
    UpdateDraftRequest,
)


def mention(**kwargs):
    return Mention(key="owner", email="owner@example.com", display_name="Owner <IT> & Co", **kwargs)


def test_mentions_escape_plain_text_and_names():
    html, kind = render_mentions("<hello>\n{{mention:owner}}", "text", [mention()])
    assert kind == "html"
    assert html.startswith("&lt;hello&gt;<br>")
    assert 'href="mailto:owner@example.com"' in html
    assert "@Owner &lt;IT&gt; &amp; Co" in html
    assert 'id="OWAAM' in html
    assert "{{mention:" not in html


def test_mention_ids_are_unique_and_html_is_preserved():
    body = "<p>{{mention:owner}} and {{mention:owner}}</p>"
    html, kind = render_mentions(body, "html", [mention()])
    from re import findall

    ids = findall(r'id="(OWAAM[A-F0-9]{32})"', html)
    assert len(set(ids)) == 2
    assert html.startswith("<p>") and html.endswith("</p>")
    assert kind == "html"


@pytest.mark.parametrize(
    "body",
    [
        '<a href="{{mention:owner}}">link</a>',
        '<a href="https://example.com">{{mention:owner}}</a>',
        "<script>{{mention:owner}}</script>",
        "<!-- {{mention:owner}} -->",
        '<p title="{{mention:owner}}">{{mention:owner}}</p>',
    ],
)
def test_mentions_cannot_be_inserted_into_markup(body):
    with pytest.raises(ValidationError):
        ReplyEmailRequest(id="source", body=body, body_type="html", mentions=[mention()])


@pytest.mark.parametrize(
    "body, mentions",
    [
        ("{{mention:owner}}", []),
        ("No placeholder", [mention()]),
        ("{{mention:owner}}", [mention(), mention()]),
        ("{{mention:unknown}}", [mention()]),
    ],
)
def test_unresolved_mentions_fail_before_any_exchange_call(body, mentions):
    with pytest.raises(ValidationError):
        DraftEmailRequest(to=["owner@example.com"], subject="test", body=body, mentions=mentions)


def test_update_mentions_require_a_body():
    with pytest.raises(ValidationError):
        UpdateDraftRequest(id="draft", mentions=[mention()])


def test_mentions_header_serializes_as_an_internet_header():
    from exchangelib import Message
    from exchangelib.version import Build, Version
    from lxml import etree

    message = Message(
        subject="test",
        body=HTMLBody("<p>mention</p>"),
        to_recipients=[Mailbox(email_address="owner@example.com")],
        x_mentions="owner@example.com,other@example.com",
    )
    xml = etree.tostring(message.to_xml(version=Version(build=Build(15, 1, 0, 0)))).decode()
    assert 'DistinguishedPropertySetId="InternetHeaders"' in xml
    assert 'PropertyName="X-Mentions"' in xml
    assert "owner@example.com,other@example.com" in xml


def test_update_body_clears_stale_mention_metadata(settings):
    backend = EWSExchangeBackend(settings)
    saves = []
    draft = SimpleNamespace(
        x_mentions="old@example.com", body="old", to_recipients=[], cc_recipients=[]
    )
    draft.save = lambda **kwargs: saves.append(kwargs)
    backend._account = SimpleNamespace(drafts=object())
    backend._fetch_item = lambda *args, **kwargs: draft
    backend.update_draft(UpdateDraftRequest(id="draft", body="No mention now"))
    assert draft.x_mentions is None
    assert saves == [{"update_fields": ["body", "x_mentions"]}]


def test_update_body_adds_native_mention_metadata_and_recipient(settings):
    backend = EWSExchangeBackend(settings)
    saves = []
    draft = SimpleNamespace(body="old", to_recipients=[], cc_recipients=[])
    draft.save = lambda **kwargs: saves.append(kwargs)
    backend._account = SimpleNamespace(drafts=object())
    backend._fetch_item = lambda *args, **kwargs: draft
    backend.update_draft(
        UpdateDraftRequest(id="draft", body="{{mention:owner}}, hello", mentions=[mention()])
    )
    assert draft.x_mentions == "owner@example.com"
    assert isinstance(draft.body, HTMLBody)
    assert [m.email_address for m in draft.to_recipients] == ["owner@example.com"]
    assert saves == [{"update_fields": ["body", "x_mentions", "to_recipients"]}]


def test_mention_header_deduplicates_and_separates_addresses(settings):
    backend = EWSExchangeBackend(settings)
    mentions = [
        mention(),
        Mention(key="second", email="Owner@example.com", display_name="Owner"),
        Mention(key="third", email="other@example.com", display_name="Other"),
    ]
    assert backend._mention_header(mentions) == "owner@example.com,other@example.com"


def test_recipient_merge_preserves_cc_and_deduplicates(settings):
    backend = EWSExchangeBackend(settings)
    item = SimpleNamespace(
        to_recipients=[Mailbox(email_address="first@example.com")],
        cc_recipients=[Mailbox(email_address="Owner@example.com")],
    )
    backend._add_recipients(
        item, ["owner@example.com", "new@example.com"], ["NEW@example.com", "cc@example.com"]
    )
    assert [m.email_address for m in item.to_recipients] == ["first@example.com", "new@example.com"]
    assert [m.email_address for m in item.cc_recipients] == ["Owner@example.com", "cc@example.com"]


def test_reply_draft_preserves_reply_all_and_attaches_without_sending(settings, tmp_path):
    settings.attachment_root = tmp_path
    attachment = tmp_path / "equipment.txt"
    attachment.write_bytes(b"serial-123")
    backend = EWSExchangeBackend(settings)
    events = []
    draft = SimpleNamespace(id="reply-draft", parent_folder_id=SimpleNamespace(id="drafts"))
    draft.attach = lambda a: events.append(("attach", a.name, a.content))
    draft.save = lambda **kwargs: events.append(("metadata", kwargs))

    class Response:
        to_recipients = [Mailbox(email_address="author@example.com")]
        cc_recipients = [Mailbox(email_address="cc@example.com")]

        def save(self, folder):
            events.append(("save", folder.id))
            return draft

        def send(self, **kwargs):
            pytest.fail("Creating a reply draft must never send")

    response = Response()

    class Source:
        subject = "Re: collect equipment"

        def create_reply_all(self, subject, body):
            events.append(("reply_all", subject, body))
            return response

    backend._account = SimpleNamespace(drafts=SimpleNamespace(id="drafts"))
    backend._fetch_item = lambda id, **kwargs: Source() if id == "source" else draft
    result = backend.create_reply_draft(
        CreateReplyDraftRequest(
            id="source",
            body="{{mention:owner}}, collect serial-123",
            mentions=[mention()],
            additional_cc=["logistics@example.com"],
            attachments=[attachment],
        )
    )
    assert result.id == "reply-draft"
    assert result.status == "draft"
    assert events[0][0:2] == ("reply_all", "Re: collect equipment")
    assert isinstance(events[0][2], HTMLBody)
    assert 'href="mailto:owner@example.com"' in events[0][2]
    assert events[1:] == [
        ("save", "drafts"),
        ("metadata", {"update_fields": ["x_mentions"]}),
        ("attach", "equipment.txt", b"serial-123"),
    ]
    assert draft.x_mentions == "owner@example.com"
    assert [m.email_address for m in response.to_recipients] == [
        "author@example.com",
        "owner@example.com",
    ]
    assert [m.email_address for m in response.cc_recipients] == [
        "cc@example.com",
        "logistics@example.com",
    ]


class SentFolder:
    id = "sent-folder"

    def __init__(self):
        self.items = []
        self.queries = []
        self.failure = None

    def filter(self, **kwargs):
        self.queries.append(kwargs)
        if self.failure:
            raise self.failure
        return self

    def only(self, *fields):
        return self

    def __getitem__(self, key):
        return self.items[key]


def delivery_backend(settings):
    sent = SentFolder()
    draft = SimpleNamespace(id="draft-id", is_draft=True, saves=[], sends=[])

    def save(**kwargs):
        draft.saves.append(kwargs)

    def send(**kwargs):
        draft.sends.append(kwargs)
        sent.items = [
            SimpleNamespace(
                id="new-sent-id", is_draft=False, datetime_sent=datetime(2026, 9, 21, tzinfo=UTC)
            )
        ]
        draft.id = None

    draft.save = save
    draft.send = send
    backend = EWSExchangeBackend(settings)
    backend._account = SimpleNamespace(
        drafts=SimpleNamespace(id="drafts"), sent=sent, primary_smtp_address="shared@example.com"
    )
    backend._fetch_item = lambda *args, **kwargs: draft
    return backend, draft, sent


def test_send_returns_real_sent_copy_not_invalid_draft_id(settings):
    backend, draft, sent = delivery_backend(settings)
    result = backend.send_draft(SendDraftRequest(id="draft-id"))
    assert result.id == "new-sent-id"
    assert result.status == "sent_confirmed"
    assert result.submission_id == draft.mcp_submission_id
    assert result.sent_copy_mailbox == "shared@example.com"
    assert result.sent_copy_folder == "sentitems"
    assert draft.saves == [{"update_fields": ["mcp_submission_id"]}]
    assert draft.sends == [{"copy_to_folder": sent}]
    assert sent.queries == [{"mcp_submission_id": result.submission_id}]


def test_no_wait_reports_submission_only(settings):
    backend, draft, sent = delivery_backend(settings)
    result = backend.send_draft(SendDraftRequest(id="draft-id", confirmation_timeout_seconds=0))
    assert result.status == "submitted"
    assert result.id is None
    assert sent.queries == []
    assert len(draft.sends) == 1


def test_verification_failure_never_turns_successful_send_into_error(settings):
    backend, draft, sent = delivery_backend(settings)
    sent.failure = RuntimeError("server read unavailable")
    result = backend.send_draft(SendDraftRequest(id="draft-id"))
    assert result.status == "submitted_unconfirmed"
    assert result.id is None
    assert "Do not resend" in result.warning
    assert len(draft.sends) == 1


def test_confirmation_waits_for_delayed_copy(settings, monkeypatch):
    backend, draft, sent = delivery_backend(settings)
    real_filter = sent.filter
    attempts = []

    def delayed_filter(**kwargs):
        attempts.append(1)
        result = real_filter(**kwargs)
        if len(attempts) == 1:
            result.items = []
        else:
            result.items = [SimpleNamespace(id="late-sent", is_draft=False, datetime_sent=None)]
        return result

    sent.filter = delayed_filter
    monkeypatch.setattr("outlook_mcp.exchange_client.email.time.sleep", lambda _: None)
    result = backend.send_draft(SendDraftRequest(id="draft-id"))
    assert result.status == "sent_confirmed"
    assert result.id == "late-sent"
    assert len(draft.sends) == 1
    assert len(attempts) == 2


def test_confirmation_timeout_does_not_resend(settings, monkeypatch):
    backend, draft, sent = delivery_backend(settings)
    draft.send = lambda **kwargs: draft.sends.append(kwargs)
    clock = iter([0, 11])
    monkeypatch.setattr("outlook_mcp.exchange_client.email.time.monotonic", lambda: next(clock))
    result = backend.send_draft(SendDraftRequest(id="draft-id"))
    assert result.status == "submitted_unconfirmed"
    assert len(draft.sends) == 1


def test_ambiguous_send_error_has_correlation_and_is_never_retried(settings):
    backend, draft, sent = delivery_backend(settings)

    def lost_response(**kwargs):
        draft.sends.append(kwargs)
        raise TimeoutError("response lost")

    draft.send = lost_response
    with pytest.raises(APIError) as exc:
        backend.send_draft(SendDraftRequest(id="draft-id"))
    assert exc.value.code == "send_outcome_unknown"
    assert exc.value.to_dict()["submission_id"] == draft.mcp_submission_id
    assert len(draft.sends) == 1


def test_stamp_failure_prevents_send(settings):
    backend, draft, sent = delivery_backend(settings)
    draft.save = lambda **kwargs: (_ for _ in ()).throw(RuntimeError("save failed"))
    with pytest.raises(APIError):
        backend.send_draft(SendDraftRequest(id="draft-id"))
    assert draft.sends == []


def test_rejects_sent_item_left_in_drafts(settings):
    backend, draft, sent = delivery_backend(settings)
    draft.is_draft = False
    with pytest.raises(APIError, match="not an unsent draft"):
        backend.send_draft(SendDraftRequest(id="draft-id"))
    assert draft.sends == [] and draft.saves == []
