from __future__ import annotations

import pytest
from pydantic import ValidationError

from outlook_mcp.models import (
    Attendee,
    ContactEmailAddress,
    CreateEventRequest,
    EmailAddress,
    MailboxInfo,
    SendEmailRequest,
)

# Exchange reports this instead of SMTP for internal senders, distribution lists
# and migrated mailboxes (routing type EX).
LEGACY_DN = (
    "/o=TANDER/ou=Exchange Administrative Group (FYDIBOHF23SPDLT)"
    "/cn=Recipients/cn=b63f03c34a-zakharov_av4"
)


@pytest.mark.parametrize(
    "model, field, extra",
    [
        (EmailAddress, "email", {}),
        (Attendee, "email", {}),
        (ContactEmailAddress, "address", {"type": "EX"}),
        (
            MailboxInfo,
            "email_address",
            {"display_name": "Захаров А.В.", "timezone": "Europe/Moscow"},
        ),
    ],
)
def test_read_models_accept_legacy_x500_address(model, field, extra) -> None:
    """A single such sender used to fail the whole folder listing."""
    assert getattr(model(**{field: LEGACY_DN}, **extra), field) == LEGACY_DN


def test_read_models_still_accept_plain_smtp() -> None:
    assert EmailAddress(email="zubov@magnit.ru").email == "zubov@magnit.ru"
    assert ContactEmailAddress(type="SMTP", address="zubov@magnit.ru").address == "zubov@magnit.ru"


def test_read_models_reject_empty_address() -> None:
    with pytest.raises(ValidationError):
        EmailAddress(email="")


def test_outgoing_recipients_stay_strict() -> None:
    """Leniency must not leak into anything that hands an address to Exchange.

    A fabricated or unroutable recipient would either bounce or, worse, deliver
    to the wrong person.
    """
    with pytest.raises(ValidationError):
        SendEmailRequest(to=[LEGACY_DN], subject="s", body="b")

    with pytest.raises(ValidationError):
        CreateEventRequest(
            subject="s",
            start="2026-07-30T10:00:00+03:00",
            end="2026-07-30T11:00:00+03:00",
            attendees=[LEGACY_DN],
        )


def test_outgoing_recipients_accept_plain_smtp() -> None:
    request = SendEmailRequest(to=["zubov@magnit.ru"], subject="s", body="b")
    assert request.to == ["zubov@magnit.ru"]
