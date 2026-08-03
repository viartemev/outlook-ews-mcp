from __future__ import annotations

from ..exchange_client import ExchangeClient
from ..models import dump_model


def ping_exchange(client: ExchangeClient, arguments: dict) -> dict:
    return dump_model(client.ping())


def get_mailbox_info(client: ExchangeClient, arguments: dict) -> dict:
    return dump_model(client.get_mailbox_info())
