from __future__ import annotations

from ..config import Settings
from .base import BaseEWSBackend
from .calendar import CalendarOperationsMixin
from .contacts import ContactOperationsMixin
from .email import EmailOperationsMixin
from .mailbox import MailboxSettingsMixin
from .protocol import ExchangeBackend


class EWSExchangeBackend(
    EmailOperationsMixin,
    MailboxSettingsMixin,
    CalendarOperationsMixin,
    ContactOperationsMixin,
    BaseEWSBackend,
):
    """Live EWS-backed implementation of ExchangeBackend, assembled from domain mixins."""


def build_default_backend(settings: Settings) -> ExchangeBackend:
    return EWSExchangeBackend(settings)
