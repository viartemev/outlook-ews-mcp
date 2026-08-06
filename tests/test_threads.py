from __future__ import annotations

from datetime import UTC, datetime
from types import SimpleNamespace

import pytest

from outlook_mcp.errors import APIError
from outlook_mcp.exchange_client import EWSExchangeBackend
from outlook_mcp.exchange_client.email import normalize_subject
from outlook_mcp.models import GetThreadRequest


def _message(item_id: str, subject: str, hour: int, body: str = "") -> SimpleNamespace:
    return SimpleNamespace(
        id=item_id,
        subject=subject,
        datetime_received=datetime(2026, 4, 7, hour, 0, tzinfo=UTC),
        conversation_id=SimpleNamespace(id="conv-1"),
        text_body=body,
        body=body,
        is_read=True,
        has_attachments=False,
        importance="Normal",
        categories=[],
        author=None,
        to_recipients=[],
        cc_recipients=[],
        bcc_recipients=[],
        attachments=[],
        headers=[],
    )


class FakeQuerySet:
    """Just enough of an exchangelib queryset for _thread_items: order_by + slicing."""

    def __init__(self, items: list) -> None:
        self.items = items

    def order_by(self, *fields) -> "FakeQuerySet":
        return self

    def __getitem__(self, key):
        return self.items[key]


class StubThreadBackend(EWSExchangeBackend):
    """Real get_thread logic with only the two EWS round trips replaced."""

    def __init__(self, settings, items_by_folder: dict[str, list], anchor=None) -> None:
        super().__init__(settings)
        self._items_by_folder = items_by_folder
        self._anchor = anchor

    def _resolve_folder(self, value):
        return value

    def _fetch_item(self, item_id, folder=None, expected_type=None):
        return self._anchor

    def _thread_items(self, folder, conversation_id, subject, limit):
        return list(self._items_by_folder.get(folder, []))


@pytest.mark.parametrize(
    ("raw", "expected"),
    [
        ("Re: Quarterly report", "Quarterly report"),
        ("RE: FW: Quarterly report", "Quarterly report"),
        ("Re[2]: Quarterly report", "Quarterly report"),
        ("Ответ: Квартальный отчёт", "Квартальный отчёт"),
        ("ПЕРЕСЛ: Квартальный отчёт", "Квартальный отчёт"),
        ("Quarterly report", "Quarterly report"),
        ("Re: Incident: disk full", "Incident: disk full"),
        ("", ""),
        (None, ""),
    ],
)
def test_normalize_subject(raw, expected) -> None:
    assert normalize_subject(raw) == expected


def test_get_thread_orders_oldest_first_and_merges_folders(settings) -> None:
    anchor = _message("email-2", "Re: Hello", 11)
    backend = StubThreadBackend(
        settings,
        {
            "inbox": [_message("email-3", "Re: Hello", 12), _message("email-1", "Hello", 10)],
            "sent": [_message("email-2", "Re: Hello", 11)],
        },
        anchor=anchor,
    )

    thread = backend.get_thread(GetThreadRequest(id="email-2"))

    assert [message.id for message in thread.messages] == ["email-1", "email-2", "email-3"]
    assert thread.conversation_id == "conv-1"
    assert thread.subject == "Hello"
    assert thread.message_count == 3
    assert thread.truncated is False


def test_get_thread_deduplicates_a_message_seen_in_two_folders(settings) -> None:
    duplicate = _message("email-1", "Hello", 10)
    backend = StubThreadBackend(
        settings,
        {"inbox": [duplicate], "archive": [_message("email-1", "Hello", 10)]},
        anchor=duplicate,
    )

    thread = backend.get_thread(GetThreadRequest(id="email-1", folders=["inbox", "archive"]))

    assert [message.id for message in thread.messages] == ["email-1"]


def test_get_thread_keeps_the_newest_messages_when_truncating(settings) -> None:
    anchor = _message("email-1", "Hello", 10)
    backend = StubThreadBackend(
        settings,
        {
            "inbox": [
                _message("email-1", "Hello", 10),
                _message("email-2", "Re: Hello", 11),
                _message("email-3", "Re: Hello", 12),
            ]
        },
        anchor=anchor,
    )

    thread = backend.get_thread(GetThreadRequest(id="email-1", folders=["inbox"], limit=2))

    assert [message.id for message in thread.messages] == ["email-2", "email-3"]
    assert thread.truncated is True
    assert thread.message_count == 2


def test_get_thread_always_contains_the_message_it_was_asked_about(settings) -> None:
    """The anchor may sit in a folder outside `folders` (Archive, a custom folder).
    Returning an empty thread for a message we are holding is simply wrong."""

    class NoMatchBackend(StubThreadBackend):
        def _thread_items(self, folder, conversation_id, subject, limit):
            return []

    anchor = _message("email-1", "Hello", 10)
    backend = NoMatchBackend(settings, {}, anchor=anchor)

    thread = backend.get_thread(GetThreadRequest(id="email-1", folders=["inbox"]))

    assert [message.id for message in thread.messages] == ["email-1"]
    assert thread.message_count == 1


def test_get_thread_requires_exactly_one_selector() -> None:
    with pytest.raises(ValueError, match="exactly one of id or conversation_id"):
        GetThreadRequest.model_validate({})

    with pytest.raises(ValueError, match="exactly one of id or conversation_id"):
        GetThreadRequest.model_validate({"id": "email-1", "conversation_id": "conv-1"})


def test_thread_items_falls_back_to_subject_when_the_restriction_is_rejected(settings) -> None:
    """Some servers reject a restriction on item:ConversationId. The fallback is
    lossy, so it only runs once that is known -- and only with a subject to use."""
    backend = EWSExchangeBackend(settings)
    attempted: list[str] = []

    class Folder:
        def filter(self, *args, **kwargs):
            if "conversation_id" in kwargs:
                attempted.append("conversation_id")
                raise ValueError("The specified restriction is not supported")
            attempted.append("subject")
            return FakeQuerySet([_message("email-1", "Hello", 10)])

    items = backend._thread_items(Folder(), "conv-1", "Re: Hello", 20)

    assert attempted == ["conversation_id", "subject"]
    assert [item.id for item in items] == ["email-1"]


def test_thread_lookup_by_conversation_id_alone_does_not_silently_return_nothing(
    settings,
) -> None:
    """With no anchor there is no subject to fall back to. Answering "this
    conversation has no messages" would be a wrong answer, not a partial one."""
    backend = EWSExchangeBackend(settings)

    class RejectingFolder:
        def filter(self, *args, **kwargs):
            raise ValueError("The specified restriction is not supported")

    with pytest.raises(APIError):
        backend._thread_items(RejectingFolder(), "conv-1", None, 20)


def test_thread_items_drops_subjects_that_merely_embed_the_normalised_one(settings) -> None:
    backend = EWSExchangeBackend(settings)
    matches = [
        _message("email-1", "Re: Hello", 10),
        _message("email-2", "Hello there, unrelated", 11),
    ]

    class Folder:
        def filter(self, *args, **kwargs):
            return FakeQuerySet(matches)

    items = backend._thread_items(Folder(), None, "Hello", 20)

    assert [item.id for item in items] == ["email-1"]
