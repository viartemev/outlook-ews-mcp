"""Per-EWS-request telemetry: the response hook, its per-call accounting, and
the counters on the tool log line."""

from __future__ import annotations

from datetime import timedelta
from types import SimpleNamespace

from outlook_mcp import telemetry
from outlook_mcp.exchange_client.base import _ews_response_hook
from outlook_mcp.server import build_registry


def _response(url: str = "https://mail.example.com/EWS/Exchange.asmx", ms: float = 250.0):
    return SimpleNamespace(
        url=url,
        status_code=200,
        elapsed=timedelta(milliseconds=ms),
    )


def test_the_hook_accumulates_into_the_current_call(caplog) -> None:
    stats = telemetry.start_collecting()

    _ews_response_hook(_response(ms=250.0))
    _ews_response_hook(_response(ms=150.0))

    assert stats.requests == 2
    assert round(stats.total_ms) == 400


def test_the_hook_is_a_no_op_outside_a_tool_call() -> None:
    """Session renewal can fire the hook from housekeeping paths that no tool
    call started; that must never crash the request that triggered it."""
    telemetry._current.set(None)

    _ews_response_hook(_response())


def test_the_hook_logs_the_path_but_never_the_query_string(caplog) -> None:
    """The query string is where anything sensitive in a URL would live."""
    telemetry.start_collecting()

    with caplog.at_level("DEBUG", logger="outlook_mcp"):
        _ews_response_hook(_response(url="https://mail.example.com/EWS/Exchange.asmx?token=s3cret"))

    joined = " ".join(record.getMessage() for record in caplog.records)
    assert "/EWS/Exchange.asmx" in joined
    assert "s3cret" not in joined


def test_each_call_starts_from_zero() -> None:
    first = telemetry.start_collecting()
    _ews_response_hook(_response())

    second = telemetry.start_collecting()

    assert first.requests == 1
    assert second.requests == 0


def test_configured_sessions_carry_the_hook_and_others_do_not(settings) -> None:
    """The hook must ride raw_session -- the one path every session the pool
    ever creates (including renewals) passes through -- and stay scoped to the
    backend that was configured, per _configure_protocol's contract."""
    from exchangelib.protocol import BaseProtocol

    from outlook_mcp.exchange_client import EWSExchangeBackend

    class FakeProtocol(BaseProtocol):
        def __init__(self):  # bypass BaseProtocol's config plumbing
            pass

    configured = FakeProtocol()
    EWSExchangeBackend(settings)._configure_protocol(configured)
    session = configured.raw_session("https://mail.example.com")
    try:
        untouched = FakeProtocol().raw_session("https://mail.example.com")
        assert _ews_response_hook in session.hooks["response"]
        assert _ews_response_hook not in untouched.hooks["response"]
    finally:
        session.close()
        untouched.close()


def test_the_tool_log_line_reports_ews_counters(settings, client, caplog) -> None:
    """duration_ms alone cannot say whether a slow call spent its time on the
    wire or in this codebase; the gap between it and ews_ms can."""
    registry = build_registry(settings=settings, client=client)

    with caplog.at_level("INFO", logger="outlook_mcp"):
        registry.call("ping_exchange", {})

    line = next(r.getMessage() for r in caplog.records if "tool=ping_exchange" in r.getMessage())
    assert "ews_requests=0" in line
    assert "ews_ms=0" in line
