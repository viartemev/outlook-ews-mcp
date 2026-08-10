import pytest
from pydantic import ValidationError

from outlook_mcp.config import Settings


def _kwargs(**overrides):
    base = {
        "EXCHANGE_SERVER": "https://mail.example.com/EWS/Exchange.asmx",
        "EXCHANGE_USERNAME": "DOMAIN\\user",
        "EXCHANGE_PASSWORD": "secret",
    }
    base.update(overrides)
    return base


def test_basic_auth_over_http_rejected_by_default():
    with pytest.raises(ValidationError, match="EXCHANGE_ALLOW_INSECURE_BASIC_AUTH"):
        Settings(
            **_kwargs(
                EXCHANGE_SERVER="http://mail.example.com/EWS/Exchange.asmx",
                EXCHANGE_AUTH_TYPE="Basic",
            )
        )


def test_basic_auth_over_http_allowed_with_explicit_opt_in():
    settings = Settings(
        **_kwargs(
            EXCHANGE_SERVER="http://mail.example.com/EWS/Exchange.asmx",
            EXCHANGE_AUTH_TYPE="Basic",
            EXCHANGE_ALLOW_INSECURE_BASIC_AUTH="true",
        )
    )
    assert settings.exchange_allow_insecure_basic_auth is True


def test_basic_auth_over_https_allowed():
    settings = Settings(
        **_kwargs(
            EXCHANGE_SERVER="https://mail.example.com/EWS/Exchange.asmx",
            EXCHANGE_AUTH_TYPE="Basic",
        )
    )
    assert settings.exchange_auth_type == "Basic"


def test_ntlm_over_http_allowed():
    settings = Settings(
        **_kwargs(
            EXCHANGE_SERVER="http://mail.example.com/EWS/Exchange.asmx",
            EXCHANGE_AUTH_TYPE="NTLM",
        )
    )
    assert settings.exchange_auth_type == "NTLM"


def test_email_body_limit_uses_exchange_prefixed_alias():
    settings = Settings(**_kwargs(EXCHANGE_EMAIL_BODY_MAX_CHARS=1234))

    assert settings.email_body_max_chars == 1234


def test_oauth2_is_rejected_until_implemented():
    with pytest.raises(ValidationError):
        Settings(**_kwargs(EXCHANGE_AUTH_TYPE="OAuth2"))


def test_mcp_max_queue_size_uses_mcp_prefixed_alias():
    settings = Settings(**_kwargs(MCP_MAX_QUEUE_SIZE=5))

    assert settings.mcp_max_queue_size == 5


def test_mcp_max_queue_size_default():
    settings = Settings(**_kwargs())

    assert settings.mcp_max_queue_size == 20


def test_attachment_limits_use_exchange_prefixed_aliases():
    settings = Settings(
        **_kwargs(
            EXCHANGE_ATTACHMENT_MAX_SIZE_MB=11,
            EXCHANGE_ATTACHMENT_MAX_COUNT=12,
            EXCHANGE_ATTACHMENT_MAX_TOTAL_SIZE_MB=26,
        )
    )

    assert settings.attachment_max_size_mb == 11
    assert settings.attachment_max_count == 12
    assert settings.attachment_max_total_size_mb == 26


@pytest.mark.parametrize(
    "env_name",
    [
        "EXCHANGE_EMAIL_ADDRESS",
        "EXCHANGE_VERSION",
        "EXCHANGE_IMPERSONATE_AS",
        "EXCHANGE_ATTACHMENT_ROOT",
        "LOG_FILE",
    ],
)
@pytest.mark.parametrize("blank", ["", "   "])
def test_blank_optional_env_values_mean_unset(env_name, blank):
    """dotenv delivers a bare `NAME=` line as an empty string, not a missing key."""
    field = {
        "EXCHANGE_EMAIL_ADDRESS": "exchange_email_address",
        "EXCHANGE_VERSION": "exchange_version",
        "EXCHANGE_IMPERSONATE_AS": "exchange_impersonate_as",
        "EXCHANGE_ATTACHMENT_ROOT": "attachment_root",
        "LOG_FILE": "log_file",
    }[env_name]

    settings = Settings(_env_file=None, **_kwargs(**{env_name: blank}))

    assert getattr(settings, field) is None


def test_blank_attachment_root_does_not_become_the_current_directory():
    """Path("") is Path("."): a blank EXCHANGE_ATTACHMENT_ROOT used to silently
    turn the server's cwd into the attachment sandbox instead of staying off."""
    settings = Settings(_env_file=None, **_kwargs(EXCHANGE_ATTACHMENT_ROOT=""))

    assert settings.attachment_root is None
