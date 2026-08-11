from __future__ import annotations

from types import SimpleNamespace

from outlook_mcp.exchange_client import EWSExchangeBackend
from outlook_mcp.models import ReplyEmailRequest


def test_signature_is_appended_per_body_type(settings) -> None:
    settings.signature_text = "-- \nIvan"
    settings.signature_html = "<p>-- Ivan</p>"
    backend = EWSExchangeBackend(settings)

    assert backend._with_signature("Hello", "text", True) == "Hello\n\n-- \nIvan"
    assert (
        backend._with_signature("<p>Hello</p>", "html", True)
        == "<p>Hello</p><br><br><p>-- Ivan</p>"
    )


def test_signature_is_not_cross_converted_between_body_types(settings) -> None:
    # Only the text signature is configured: an html body gets nothing rather
    # than raw text pasted into markup.
    settings.signature_text = "-- Ivan"
    settings.signature_html = None
    backend = EWSExchangeBackend(settings)

    assert backend._with_signature("<p>Hello</p>", "html", True) == "<p>Hello</p>"


def test_signature_respects_include_signature_false(settings) -> None:
    settings.signature_text = "-- Ivan"
    backend = EWSExchangeBackend(settings)

    assert backend._with_signature("Hello", "text", False) == "Hello"


def test_signature_absent_when_not_configured(settings) -> None:
    backend = EWSExchangeBackend(settings)

    assert backend._with_signature("Hello", "text", True) == "Hello"


def test_reply_email_appends_the_text_signature(settings) -> None:
    settings.signature_text = "-- Ivan"
    backend = EWSExchangeBackend(settings)
    captured: list[str] = []

    class FakeResponse:
        def send(self):
            pass

    class FakeItem:
        subject = "Hello"

        def create_reply(self, subject, body):
            captured.append(body)
            return FakeResponse()

    def fetch(ids, folder=None):
        yield FakeItem()

    backend._account = SimpleNamespace(fetch=fetch)

    backend.reply_email(ReplyEmailRequest(id="msg-1", body="Reply body"))

    assert captured == ["Reply body\n\n-- Ivan"]


def test_reply_email_can_opt_out_of_the_signature(settings) -> None:
    settings.signature_text = "-- Ivan"
    backend = EWSExchangeBackend(settings)
    captured: list[str] = []

    class FakeResponse:
        def send(self):
            pass

    class FakeItem:
        subject = "Hello"

        def create_reply(self, subject, body):
            captured.append(body)
            return FakeResponse()

    def fetch(ids, folder=None):
        yield FakeItem()

    backend._account = SimpleNamespace(fetch=fetch)

    backend.reply_email(ReplyEmailRequest(id="msg-1", body="Reply body", include_signature=False))

    assert captured == ["Reply body"]
