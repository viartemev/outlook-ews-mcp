from __future__ import annotations

from dataclasses import dataclass, field
from typing import Literal

from .config import Settings
from .errors import APIError


@dataclass(slots=True)
class AuthContext:
    auth_type: Literal["NTLM", "Basic"]
    username: str
    password: str = field(repr=False)
    primary_smtp_address: str
    impersonate_as: str | None = None


def build_auth_context(settings: Settings) -> AuthContext:
    primary_smtp_address = (
        settings.exchange_impersonate_as
        or settings.exchange_email_address
        or (settings.exchange_username if "@" in settings.exchange_username else None)
    )
    if not primary_smtp_address:
        raise APIError(
            "validation_error",
            "unable to determine mailbox SMTP address",
            details=[
                {
                    "field": "EXCHANGE_EMAIL_ADDRESS",
                    "reason": "set this when EXCHANGE_USERNAME is not an email address",
                }
            ],
        )
    plaintext_password = settings.exchange_password.get_secret_value()
    return AuthContext(
        auth_type=settings.exchange_auth_type,
        username=settings.exchange_username,
        password=plaintext_password,
        primary_smtp_address=primary_smtp_address,
        impersonate_as=settings.exchange_impersonate_as,
    )
