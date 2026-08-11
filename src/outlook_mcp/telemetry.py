"""Per-tool-call accounting of the EWS requests made underneath it.

The tool log line in server.py reports duration_ms for the whole call; that
number alone cannot say whether a slow call spent its time on the wire or in
this codebase. The response hook installed by _configure_protocol
(exchange_client/base.py) reports every EWS HTTP request here, and
ToolRegistry.call folds the totals into the same log line -- the gap between
duration_ms and ews_ms is ours, the rest is Exchange and the network.

A ContextVar carries the accumulator: ToolRegistry.call starts a fresh one,
and the hook fires on the same thread (and therefore in the same context copy)
that issued the HTTP request, so concurrent tool calls never mix their counts.
"""

from __future__ import annotations

from contextvars import ContextVar
from dataclasses import dataclass


@dataclass
class EwsCallStats:
    requests: int = 0
    total_ms: float = 0.0


_current: ContextVar[EwsCallStats | None] = ContextVar("ews_call_stats", default=None)


def start_collecting() -> EwsCallStats:
    """Begin a fresh accumulator for the current call and return it."""
    stats = EwsCallStats()
    _current.set(stats)
    return stats


def record_request(elapsed_ms: float) -> None:
    """Count one EWS HTTP request; a no-op outside a collecting tool call."""
    stats = _current.get()
    if stats is not None:
        stats.requests += 1
        stats.total_ms += elapsed_ms
