from __future__ import annotations

from outlook_mcp.config import Settings
from outlook_mcp.errors import APIError
from outlook_mcp.exchange_client import ExchangeClient
from outlook_mcp.models import PingResult, SendResult
import pytest


def _settings(**overrides) -> Settings:
    payload = {
        "EXCHANGE_SERVER": "https://mail.example.com/EWS/Exchange.asmx",
        "EXCHANGE_USERNAME": "user@example.com",
        "EXCHANGE_PASSWORD": "secret",
    }
    payload.update(overrides)
    return Settings(**payload)  # type: ignore[arg-type]


def _busy_error() -> APIError:
    # Same shape _map_exception produces for ErrorServerBusy et al under FailFast.
    return APIError("exchange_unavailable", "exchange reported itself busy", retryable=True)


class _FlakyBackend:
    """Fails a fixed number of times with a retryable error, then succeeds."""

    def __init__(self, failures: int) -> None:
        self.failures = failures
        self.calls = 0

    def ping(self) -> PingResult:
        self.calls += 1
        if self.calls <= self.failures:
            raise _busy_error()
        return PingResult(status="ok", server="mail.example.com", version="2019", latency_ms=1)

    def send_email(self, request) -> SendResult:  # noqa: ANN001
        self.calls += 1
        raise _busy_error()


class _AlwaysFailsNonRetryable:
    def __init__(self) -> None:
        self.calls = 0

    def ping(self) -> PingResult:
        self.calls += 1
        raise APIError("validation_error", "bad request", retryable=False)


def test_retry_read_retries_a_busy_read_until_it_succeeds(monkeypatch) -> None:
    monkeypatch.setattr(ExchangeClient, "_RETRY_BASE_SECONDS", 0.001)
    backend = _FlakyBackend(failures=2)
    client = ExchangeClient(
        settings=_settings(EXCHANGE_TIMEOUT=1, EXCHANGE_MAX_RETRY_WAIT_SECONDS=10),
        backend=backend,
    )

    result = client.ping()

    assert result.status == "ok"
    assert backend.calls == 3


def test_retry_read_gives_up_once_the_deadline_is_exhausted() -> None:
    backend = _FlakyBackend(failures=999)
    client = ExchangeClient(settings=_settings(EXCHANGE_MAX_RETRY_WAIT_SECONDS=0), backend=backend)

    try:
        client.ping()
    except APIError as exc:
        assert exc.code == "exchange_unavailable"
    else:  # pragma: no cover
        raise AssertionError("expected the retryable error to propagate")

    # A zero-second budget means exactly one attempt, not an unbounded retry loop.
    assert backend.calls == 1


def test_retry_read_does_not_retry_non_retryable_errors() -> None:
    backend = _AlwaysFailsNonRetryable()
    client = ExchangeClient(settings=_settings(EXCHANGE_MAX_RETRY_WAIT_SECONDS=5), backend=backend)

    try:
        client.ping()
    except APIError as exc:
        assert exc.code == "validation_error"
    else:  # pragma: no cover
        raise AssertionError("expected the non-retryable error to propagate immediately")

    assert backend.calls == 1


def test_writes_are_never_auto_retried() -> None:
    """A retryable failure on send/create/delete must surface immediately: retrying
    it could repeat a side effect that already happened on the server before the
    error came back."""
    from outlook_mcp.models import SendEmailRequest

    backend = _FlakyBackend(failures=999)
    client = ExchangeClient(settings=_settings(EXCHANGE_MAX_RETRY_WAIT_SECONDS=5), backend=backend)
    request = SendEmailRequest.model_validate(
        {"to": ["user@example.com"], "subject": "hi", "body": "hi"}
    )

    try:
        client.send_email(request)
    except APIError:
        pass
    else:  # pragma: no cover
        raise AssertionError("expected the write to propagate its error immediately")

    assert backend.calls == 1


def test_retry_does_not_start_an_attempt_without_a_full_timeout_remaining(monkeypatch) -> None:
    clock = [0.0]

    class SlowBusyBackend:
        calls = 0

        def ping(self):
            self.calls += 1
            clock[0] += 4
            raise _busy_error()

    monkeypatch.setattr("outlook_mcp.exchange_client.facade.time.monotonic", lambda: clock[0])
    monkeypatch.setattr(
        "outlook_mcp.exchange_client.facade.time.sleep",
        lambda seconds: clock.__setitem__(0, clock[0] + seconds),
    )
    backend = SlowBusyBackend()
    client = ExchangeClient(
        settings=_settings(EXCHANGE_TIMEOUT=4, EXCHANGE_MAX_RETRY_WAIT_SECONDS=5),
        backend=backend,
    )

    with pytest.raises(APIError):
        client.ping()

    assert backend.calls == 1
    assert clock[0] <= 9
