from __future__ import annotations

import json
import os
from datetime import datetime, timedelta
from typing import Any
from zoneinfo import ZoneInfo

from .config import get_settings
from .exchange_client import ExchangeClient, build_default_backend
from .models import ListEmailsRequest, ListEventsRequest


def _env_flag(name: str) -> bool:
    return os.getenv(name, "").strip().lower() in {"1", "true", "yes", "on"}


def _mask_email(value: str | None) -> str | None:
    if not value or "@" not in value:
        return value
    local, domain = value.split("@", 1)
    if len(local) <= 2:
        masked_local = local[:1] + "*" * max(len(local) - 1, 0)
    else:
        masked_local = local[:2] + "*" * (len(local) - 2)
    return f"{masked_local}@{domain}"


def _sanitize_mailbox(payload: dict[str, Any]) -> dict[str, Any]:
    sanitized = dict(payload)
    sanitized["email_address"] = _mask_email(payload.get("email_address"))
    return sanitized


def main() -> None:
    settings = get_settings()
    backend = build_default_backend(settings)
    client = ExchangeClient(settings=settings, backend=backend)

    ping = client.ping().model_dump(mode="json")
    mailbox = client.get_mailbox_info().model_dump(mode="json")

    start = datetime.now(ZoneInfo(mailbox["timezone"]))
    end = start + timedelta(days=7)
    inbox = [
        item.model_dump(mode="json", by_alias=True)
        for item in client.list_emails(ListEmailsRequest(limit=5))
    ]
    events = [
        item.model_dump(mode="json", by_alias=True)
        for item in client.list_events(ListEventsRequest(start=start, end=end))
    ]

    include_data = _env_flag("OUTLOOK_MCP_SMOKE_INCLUDE_DATA")
    payload: dict[str, Any] = {
        "ping_exchange": ping,
        "mailbox_info": _sanitize_mailbox(mailbox),
        "recent_inbox_count": len(inbox),
        "upcoming_events_count": len(events),
        "data_included": include_data,
    }
    if include_data:
        payload["recent_inbox"] = inbox
        payload["upcoming_events"] = events

    print(json.dumps(payload, ensure_ascii=False, indent=2))
