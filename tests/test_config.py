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


def test_email_mime_limit_uses_exchange_prefixed_alias():
    settings = Settings(**_kwargs(EXCHANGE_EMAIL_MIME_MAX_SIZE_MB=12))

    assert settings.email_mime_max_size_mb == 12


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
        "EXCHANGE_CA_BUNDLE",
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
        "EXCHANGE_CA_BUNDLE": "exchange_ca_bundle",
        "LOG_FILE": "log_file",
    }[env_name]

    settings = Settings(_env_file=None, **_kwargs(**{env_name: blank}))

    assert getattr(settings, field) is None


def test_blank_attachment_root_does_not_become_the_current_directory():
    """Path("") is Path("."): a blank EXCHANGE_ATTACHMENT_ROOT used to silently
    turn the server's cwd into the attachment sandbox instead of staying off."""
    settings = Settings(_env_file=None, **_kwargs(EXCHANGE_ATTACHMENT_ROOT=""))

    assert settings.attachment_root is None


def test_attachment_root_must_be_absolute():
    with pytest.raises(ValidationError, match="must be an absolute path"):
        Settings(_env_file=None, **_kwargs(EXCHANGE_ATTACHMENT_ROOT="relative/path"))


def test_tls_verify_defaults_to_system_store():
    settings = Settings(_env_file=None, **_kwargs())

    assert settings.exchange_tls_verify is True


def test_tls_verify_uses_ca_bundle_when_set(tmp_path):
    bundle = tmp_path / "ca.pem"
    bundle.write_text("dummy")

    settings = Settings(_env_file=None, **_kwargs(EXCHANGE_CA_BUNDLE=str(bundle)))

    assert settings.exchange_tls_verify == str(bundle)


def test_tls_verify_disabled_ignores_ca_bundle(tmp_path):
    bundle = tmp_path / "ca.pem"
    bundle.write_text("dummy")

    settings = Settings(
        _env_file=None,
        **_kwargs(EXCHANGE_CA_BUNDLE=str(bundle), EXCHANGE_VERIFY_SSL="false"),
    )

    assert settings.exchange_tls_verify is False


def test_missing_ca_bundle_is_rejected(tmp_path):
    with pytest.raises(ValidationError, match="EXCHANGE_CA_BUNDLE file not found"):
        Settings(_env_file=None, **_kwargs(EXCHANGE_CA_BUNDLE=str(tmp_path / "nope.pem")))


@pytest.mark.parametrize("transport", ["streamable-http", "sse"])
def test_http_transports_require_an_auth_token(transport):
    with pytest.raises(ValidationError, match="MCP_HTTP_AUTH_TOKEN"):
        Settings(**_kwargs(MCP_TRANSPORT=transport))


def test_blank_http_auth_token_counts_as_missing():
    # .env.example ships `MCP_HTTP_AUTH_TOKEN=`; an empty token must not pass.
    with pytest.raises(ValidationError, match="MCP_HTTP_AUTH_TOKEN"):
        Settings(**_kwargs(MCP_TRANSPORT="streamable-http", MCP_HTTP_AUTH_TOKEN=""))


def test_stdio_transport_needs_no_auth_token():
    assert Settings(**_kwargs(MCP_TRANSPORT="stdio")).mcp_http_auth_token is None
