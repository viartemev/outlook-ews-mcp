from __future__ import annotations

import base64
from types import SimpleNamespace

import pytest

from outlook_mcp.errors import APIError
from outlook_mcp.exchange_client import EWSExchangeBackend
from outlook_mcp.models import GetEmailMimeRequest


def _backend_with_item(item: SimpleNamespace) -> EWSExchangeBackend:
    backend = EWSExchangeBackend.__new__(EWSExchangeBackend)
    backend.settings = SimpleNamespace(exchange_timezone_fallback=None)

    def fetch(ids, folder=None):
        yield item

    backend._account = SimpleNamespace(fetch=fetch)
    return backend


def test_get_email_mime_returns_base64_content() -> None:
    item = SimpleNamespace(subject="Hello world", mime_content=b"From: a@b.com\r\n\r\nhi")
    backend = _backend_with_item(item)

    result = backend.get_email_mime(GetEmailMimeRequest(id="email-1"))

    assert result.id == "email-1"
    assert result.filename == "Hello world.eml"
    assert result.content_type == "message/rfc822"
    assert result.size == len(item.mime_content)
    assert base64.b64decode(result.mime_base64) == item.mime_content


def test_get_email_mime_falls_back_to_generic_filename_without_subject() -> None:
    item = SimpleNamespace(subject=None, mime_content=b"raw")
    backend = _backend_with_item(item)

    result = backend.get_email_mime(GetEmailMimeRequest(id="email-1"))

    assert result.filename == "message.eml"


def test_get_email_mime_raises_when_exchange_returns_no_content() -> None:
    item = SimpleNamespace(subject="Hello", mime_content=None)
    backend = _backend_with_item(item)

    with pytest.raises(APIError):
        backend.get_email_mime(GetEmailMimeRequest(id="email-1"))
