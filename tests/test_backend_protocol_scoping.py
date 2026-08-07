from __future__ import annotations

import threading
import time
from unittest.mock import patch

import pytest
from exchangelib.ewsdatetime import EWSTimeZone

from outlook_mcp.config import Settings
from outlook_mcp.errors import APIError
from outlook_mcp.exchange_client import EWSExchangeBackend

_UNKNOWN_TIMEZONE_GUID = "2a2b3c4d-0000-0000-0000-000000000001"


def _settings(**overrides) -> Settings:
    """Builds Settings with no on-disk .env fallback, so these tests can't accidentally pass by
    picking up a developer's local EXCHANGE_EMAIL_ADDRESS/etc rather than exercising the values
    given here (see auth.build_auth_context, which needs an '@' username or an explicit email)."""
    return Settings(
        _env_file=None,
        EXCHANGE_SERVER=overrides.pop(
            "EXCHANGE_SERVER", "https://mail.example.com/EWS/Exchange.asmx"
        ),
        EXCHANGE_USERNAME=overrides.pop("EXCHANGE_USERNAME", "user@example.com"),
        EXCHANGE_PASSWORD=overrides.pop("EXCHANGE_PASSWORD", "secret"),
        **overrides,
    )


def test_two_backends_keep_their_own_ssl_verification_and_timeout() -> None:
    """A second backend used to silently inherit the first backend's process-global
    BaseProtocol.raw_session/TIMEOUT patch. Distinct servers/credentials get exchangelib's
    own separate (uncached) Protocol instances, so each backend's settings must stick."""
    verifying = EWSExchangeBackend(
        _settings(
            EXCHANGE_SERVER="https://mail1.example.com/EWS/Exchange.asmx", EXCHANGE_TIMEOUT=30
        )
    )
    non_verifying = EWSExchangeBackend(
        _settings(
            EXCHANGE_SERVER="https://mail2.example.com/EWS/Exchange.asmx",
            EXCHANGE_USERNAME="other-user@example.com",
            EXCHANGE_VERIFY_SSL=False,
            EXCHANGE_TIMEOUT=5,
        )
    )

    verifying_protocol = verifying.account.protocol
    non_verifying_protocol = non_verifying.account.protocol

    assert verifying_protocol is not non_verifying_protocol
    assert verifying_protocol.raw_session("https://mail1.example.com").verify is True
    assert non_verifying_protocol.raw_session("https://mail2.example.com").verify is False
    assert verifying_protocol.TIMEOUT == 30
    assert non_verifying_protocol.TIMEOUT == 5


def test_timezone_fallback_follows_the_most_recently_active_backend() -> None:
    """EWSTimeZone.from_ms_id is patched once, process-wide, since exchangelib gives it no way
    to know which account triggered the parse. The fallback timezone must still track whichever
    backend's `.account` was most recently accessed, not whichever backend was built first."""
    moscow_backend = EWSExchangeBackend(
        _settings(
            EXCHANGE_SERVER="https://mail1.example.com/EWS/Exchange.asmx",
            EXCHANGE_TIMEZONE_FALLBACK="Europe/Moscow",
        )
    )
    new_york_backend = EWSExchangeBackend(
        _settings(
            EXCHANGE_SERVER="https://mail2.example.com/EWS/Exchange.asmx",
            EXCHANGE_USERNAME="other-user@example.com",
            EXCHANGE_TIMEZONE_FALLBACK="America/New_York",
        )
    )

    _ = moscow_backend.account
    assert EWSTimeZone.from_ms_id(_UNKNOWN_TIMEZONE_GUID).key == "Europe/Moscow"

    _ = new_york_backend.account
    assert EWSTimeZone.from_ms_id(_UNKNOWN_TIMEZONE_GUID).key == "America/New_York"


def test_exchange_version_setting_configures_account_protocol_version() -> None:
    """EXCHANGE_VERSION used to be parsed and stashed but never actually reach
    exchangelib.Configuration, so it silently did nothing."""
    backend = EWSExchangeBackend(
        _settings(
            EXCHANGE_SERVER="https://mail3.example.com/EWS/Exchange.asmx",
            EXCHANGE_USERNAME="version-user@example.com",
            EXCHANGE_VERSION="EXCHANGE_2016",
        )
    )

    assert backend.account.protocol.config.version.api_version == "Exchange2016"


def test_unset_exchange_version_leaves_autodetection_untouched() -> None:
    backend = EWSExchangeBackend(
        _settings(
            EXCHANGE_SERVER="https://mail4.example.com/EWS/Exchange.asmx",
            EXCHANGE_USERNAME="unset-version-user@example.com",
        )
    )

    assert backend.account.protocol.config.version is None


def test_concurrent_account_access_builds_the_account_once() -> None:
    """ToolGateway runs tool bodies in a thread pool, so once MCP_MAX_CONCURRENCY > 1 the
    first burst of concurrent calls can all reach `.account` before any of them has finished
    building it. Without a lock, each would see self._account as None and race to build its
    own Account/Configuration/protocol; the lock in the `account` property must collapse that
    down to a single build shared by every caller."""
    backend = EWSExchangeBackend(_settings())
    real_build_account = EWSExchangeBackend._build_account
    call_count = 0
    call_count_lock = threading.Lock()

    def slow_build_account(self):
        nonlocal call_count
        with call_count_lock:
            call_count += 1
        # Widens the race window so concurrent callers overlap inside _build_account
        # rather than happening to interleave around it.
        time.sleep(0.05)
        return real_build_account(self)

    results: list[object] = []
    with patch.object(EWSExchangeBackend, "_build_account", slow_build_account):
        threads = [
            threading.Thread(target=lambda: results.append(backend.account)) for _ in range(16)
        ]
        for thread in threads:
            thread.start()
        for thread in threads:
            thread.join()

    assert call_count == 1
    assert len(results) == 16
    assert len({id(result) for result in results}) == 1


def test_unknown_exchange_version_raises_validation_error() -> None:
    backend = EWSExchangeBackend(
        _settings(
            EXCHANGE_SERVER="https://mail5.example.com/EWS/Exchange.asmx",
            EXCHANGE_USERNAME="bad-version-user@example.com",
            EXCHANGE_VERSION="NOT_A_REAL_VERSION",
        )
    )

    with pytest.raises(APIError) as excinfo:
        _ = backend.account

    assert excinfo.value.code == "validation_error"
