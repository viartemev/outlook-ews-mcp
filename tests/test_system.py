from __future__ import annotations

from outlook_mcp.tools.system import get_mailbox_info, ping_exchange


def test_ping_exchange(client) -> None:
    result = ping_exchange(client, {})
    assert result["status"] == "ok"
    assert result["latency_ms"] == 42


def test_get_mailbox_info(client) -> None:
    result = get_mailbox_info(client, {})
    assert result["email_address"] == "user@example.com"
    assert result["display_name"] == "Test User"
