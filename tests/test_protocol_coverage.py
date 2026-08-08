from __future__ import annotations

import inspect
from typing import get_type_hints

import pytest

from outlook_mcp import models
from outlook_mcp.errors import ExchangeUnavailableError
from outlook_mcp.exchange_client import ExchangeClient
from outlook_mcp.exchange_client.protocol import ExchangeBackend
from outlook_mcp.exchange_client.unconfigured import UnconfiguredExchangeBackend

PROTOCOL_METHODS = sorted(
    name
    for name, member in vars(ExchangeBackend).items()
    if not name.startswith("_") and callable(member)
)


def _sample_request(method_name: str):
    """Build the request object a protocol method expects, from its type hints."""
    hints = get_type_hints(getattr(ExchangeBackend, method_name))
    request_type = hints.get("request")
    if request_type is None:
        return None
    samples: dict[type, object] = {
        models.ListEmailsRequest: models.ListEmailsRequest(),
        models.GetEmailRequest: models.GetEmailRequest(id="email-1"),
        models.GetThreadRequest: models.GetThreadRequest(id="email-1"),
        models.SearchEmailsRequest: models.SearchEmailsRequest(query="hello"),
        models.SendEmailRequest: models.SendEmailRequest(
            to=["user@example.com"], subject="Hi", body="Hello"
        ),
        models.ReplyEmailRequest: models.ReplyEmailRequest(id="email-1", body="Hi"),
        models.ForwardEmailRequest: models.ForwardEmailRequest(
            id="email-1", to=["user@example.com"]
        ),
        models.FolderActionRequest: models.FolderActionRequest(id="email-1", folder="inbox"),
        models.DeleteEmailRequest: models.DeleteEmailRequest(id="email-1"),
        models.MarkEmailRequest: models.MarkEmailRequest(id="email-1"),
        models.CategorizeEmailRequest: models.CategorizeEmailRequest(id="email-1"),
        models.ListCategoriesRequest: models.ListCategoriesRequest(),
        models.ListFoldersRequest: models.ListFoldersRequest(),
        models.CreateFolderRequest: models.CreateFolderRequest(name="Projects"),
        models.RenameFolderRequest: models.RenameFolderRequest(folder="Projects", name="Archive2"),
        models.DeleteFolderRequest: models.DeleteFolderRequest(folder="Projects"),
        models.DraftEmailRequest: models.DraftEmailRequest(
            to=["user@example.com"], subject="Hi", body="Hello"
        ),
        models.SendDraftRequest: models.SendDraftRequest(id="draft-1"),
        models.GetAttachmentRequest: models.GetAttachmentRequest(
            email_id="email-1", attachment_id="att-1"
        ),
        models.AddAttachmentRequest: models.AddAttachmentRequest(
            email_id="email-1", path="/tmp/x.txt"
        ),
        models.DeleteAttachmentRequest: models.DeleteAttachmentRequest(
            email_id="email-1", attachment_id="att-1"
        ),
        models.BulkMoveEmailsRequest: models.BulkMoveEmailsRequest(ids=["email-1"], folder="inbox"),
        models.BulkDeleteEmailsRequest: models.BulkDeleteEmailsRequest(ids=["email-1"]),
        models.ListEventsRequest: models.ListEventsRequest(
            start="2026-04-13T09:00:00+00:00", end="2026-04-13T18:00:00+00:00"
        ),
        models.GetEventRequest: models.GetEventRequest(id="event-1"),
        models.CreateEventRequest: models.CreateEventRequest(
            subject="Sync", start="2026-04-13T09:00:00+00:00", end="2026-04-13T10:00:00+00:00"
        ),
        models.UpdateEventRequest: models.UpdateEventRequest(id="event-1", subject="Sync"),
        models.DeleteEventRequest: models.DeleteEventRequest(id="event-1"),
        models.RespondToInviteRequest: models.RespondToInviteRequest(
            id="event-1", response="accept"
        ),
        models.FindFreeSlotsRequest: models.FindFreeSlotsRequest(
            attendees=["user@example.com"],
            duration=60,
            start="2026-04-13T09:00:00+00:00",
            end="2026-04-13T18:00:00+00:00",
        ),
        models.SearchContactsRequest: models.SearchContactsRequest(query="ivan"),
        models.GetContactRequest: models.GetContactRequest(id="contact-1"),
        models.CreateContactRequest: models.CreateContactRequest(display_name="Ivan"),
        models.UpdateContactRequest: models.UpdateContactRequest(
            id="contact-1", display_name="Ivan"
        ),
        models.DeleteContactRequest: models.DeleteContactRequest(id="contact-1"),
        models.CreateInboxRuleRequest: models.CreateInboxRuleRequest(
            display_name="r",
            conditions=models.CreateInboxRuleConditions(has_attachments=True),
            actions=models.CreateInboxRuleActions(mark_as_read=True),
        ),
        models.UpdateInboxRuleRequest: models.UpdateInboxRuleRequest(id="r-1", is_enabled=False),
        models.DeleteInboxRuleRequest: models.DeleteInboxRuleRequest(id="r-1"),
        models.OutOfOfficeSettings: models.OutOfOfficeSettings(state="disabled"),
    }
    assert request_type in samples, f"no sample request for {method_name}: {request_type}"
    return samples[request_type]


def _invoke(target, method_name: str):
    method = getattr(target, method_name)
    request = _sample_request(method_name)
    return method() if request is None else method(request)


@pytest.mark.parametrize("method_name", PROTOCOL_METHODS)
def test_the_unconfigured_backend_refuses_every_method(settings, method_name) -> None:
    """Every protocol method must fail the same, actionable way when no real
    backend is wired -- not with an AttributeError from a missing stub."""
    backend = UnconfiguredExchangeBackend(settings)

    with pytest.raises(ExchangeUnavailableError):
        _invoke(backend, method_name)


@pytest.mark.parametrize("method_name", PROTOCOL_METHODS)
def test_the_facade_delegates_every_method_to_the_backend(settings, method_name) -> None:
    calls: list[tuple[str, object]] = []

    class RecordingBackend:
        def __getattr__(self, name):
            def record(request=None):
                calls.append((name, request))
                return None

            return record

    client = ExchangeClient(settings=settings, backend=RecordingBackend())

    _invoke(client, method_name)

    assert [name for name, _ in calls] == [method_name]


def test_every_protocol_method_has_a_facade_counterpart() -> None:
    for method_name in PROTOCOL_METHODS:
        facade_method = getattr(ExchangeClient, method_name, None)
        assert callable(facade_method), f"ExchangeClient is missing {method_name}()"
        expected = inspect.signature(getattr(ExchangeBackend, method_name))
        actual = inspect.signature(facade_method)
        assert list(actual.parameters) == list(expected.parameters), method_name
