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


def test_sse_transport_without_auth_token_rejected():
    with pytest.raises(ValidationError, match="MCP_SSE_AUTH_TOKEN"):
        Settings(**_kwargs(MCP_TRANSPORT="sse"))


def test_sse_transport_with_auth_token_allowed():
    settings = Settings(**_kwargs(MCP_TRANSPORT="sse", MCP_SSE_AUTH_TOKEN="s3cret"))
    assert settings.mcp_sse_auth_token == "s3cret"


def test_stdio_transport_does_not_require_auth_token():
    settings = Settings(**_kwargs(MCP_TRANSPORT="stdio"))
    assert settings.mcp_sse_auth_token is None
