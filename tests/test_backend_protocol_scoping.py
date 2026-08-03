from __future__ import annotations

from exchangelib.ewsdatetime import EWSTimeZone

from outlook_mcp.config import Settings
from outlook_mcp.exchange_client import EWSExchangeBackend

_UNKNOWN_TIMEZONE_GUID = "2a2b3c4d-0000-0000-0000-000000000001"


def _settings(**overrides) -> Settings:
    """Builds Settings with no on-disk .env fallback, so these tests can't accidentally pass by
    picking up a developer's local EXCHANGE_EMAIL_ADDRESS/etc rather than exercising the values
    given here (see auth.build_auth_context, which needs an '@' username or an explicit email)."""
    return Settings(
        _env_file=None,
        EXCHANGE_SERVER=overrides.pop("EXCHANGE_SERVER", "https://mail.example.com/EWS/Exchange.asmx"),
        EXCHANGE_USERNAME=overrides.pop("EXCHANGE_USERNAME", "user@example.com"),
        EXCHANGE_PASSWORD=overrides.pop("EXCHANGE_PASSWORD", "secret"),
        **overrides,
    )


def test_two_backends_keep_their_own_ssl_verification_and_timeout() -> None:
    """A second backend used to silently inherit the first backend's process-global
    BaseProtocol.raw_session/TIMEOUT patch. Distinct servers/credentials get exchangelib's
    own separate (uncached) Protocol instances, so each backend's settings must stick."""
    verifying = EWSExchangeBackend(
        _settings(EXCHANGE_SERVER="https://mail1.example.com/EWS/Exchange.asmx", EXCHANGE_TIMEOUT=30)
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
        _settings(EXCHANGE_SERVER="https://mail1.example.com/EWS/Exchange.asmx", EXCHANGE_TIMEZONE="Europe/Moscow")
    )
    new_york_backend = EWSExchangeBackend(
        _settings(
            EXCHANGE_SERVER="https://mail2.example.com/EWS/Exchange.asmx",
            EXCHANGE_USERNAME="other-user@example.com",
            EXCHANGE_TIMEZONE="America/New_York",
        )
    )

    _ = moscow_backend.account
    assert EWSTimeZone.from_ms_id(_UNKNOWN_TIMEZONE_GUID).key == "Europe/Moscow"

    _ = new_york_backend.account
    assert EWSTimeZone.from_ms_id(_UNKNOWN_TIMEZONE_GUID).key == "America/New_York"
