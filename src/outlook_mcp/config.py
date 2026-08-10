from __future__ import annotations

from functools import lru_cache
from pathlib import Path
from typing import Literal
from urllib.parse import urlparse

from pydantic import Field, SecretStr, field_validator, model_validator
from pydantic_settings import BaseSettings, SettingsConfigDict


class Settings(BaseSettings):
    model_config = SettingsConfigDict(
        env_file=".env",
        env_file_encoding="utf-8",
        case_sensitive=False,
        extra="ignore",
    )

    exchange_server: str = Field(alias="EXCHANGE_SERVER")
    exchange_username: str = Field(alias="EXCHANGE_USERNAME")
    exchange_password: SecretStr = Field(alias="EXCHANGE_PASSWORD")
    exchange_email_address: str | None = Field(default=None, alias="EXCHANGE_EMAIL_ADDRESS")
    exchange_verify_ssl: bool = Field(default=True, alias="EXCHANGE_VERIFY_SSL")
    exchange_auth_type: Literal["NTLM", "Basic"] = Field(
        default="NTLM",
        alias="EXCHANGE_AUTH_TYPE",
    )
    exchange_allow_insecure_basic_auth: bool = Field(
        default=False, alias="EXCHANGE_ALLOW_INSECURE_BASIC_AUTH"
    )
    exchange_version: str | None = Field(default=None, alias="EXCHANGE_VERSION")
    exchange_timeout: int = Field(default=30, alias="EXCHANGE_TIMEOUT", ge=1, le=300)
    #: Total wall-clock budget for ExchangeClient's own retry of read-only calls
    #: when Exchange reports itself busy (0 disables retries). Not passed to
    #: exchangelib's retry_policy, which is always FailFast -- see
    #: exchange_client/base.py and ExchangeClient._retry_read.
    exchange_max_retry_wait: int = Field(
        default=90, alias="EXCHANGE_MAX_RETRY_WAIT_SECONDS", ge=0, le=3600
    )
    # Only used as a fallback when Exchange reports a timezone as an unresolvable GUID
    # (not applied as a general default -- naive datetimes and all-day math use
    # account.default_timezone instead, see exchange_client/base.py).
    exchange_timezone_fallback: str = Field(
        default="Europe/Moscow", alias="EXCHANGE_TIMEZONE_FALLBACK"
    )
    exchange_impersonate_as: str | None = Field(default=None, alias="EXCHANGE_IMPERSONATE_AS")
    attachment_max_size_mb: int = Field(
        default=10, alias="EXCHANGE_ATTACHMENT_MAX_SIZE_MB", ge=1, le=100
    )
    attachment_max_count: int = Field(
        default=10, alias="EXCHANGE_ATTACHMENT_MAX_COUNT", ge=1, le=100
    )
    attachment_max_total_size_mb: int = Field(
        default=25, alias="EXCHANGE_ATTACHMENT_MAX_TOTAL_SIZE_MB", ge=1, le=500
    )
    attachment_root: Path | None = Field(default=None, alias="EXCHANGE_ATTACHMENT_ROOT")
    email_body_max_chars: int = Field(
        default=200_000, alias="EXCHANGE_EMAIL_BODY_MAX_CHARS", ge=1_000, le=5_000_000
    )

    mcp_transport: Literal["stdio", "sse"] = Field(default="stdio", alias="MCP_TRANSPORT")
    mcp_sse_host: str = Field(default="127.0.0.1", alias="MCP_SSE_HOST")
    mcp_sse_port: int = Field(default=8080, alias="MCP_SSE_PORT", ge=1, le=65535)
    #: How many tool calls execute at once; more calls than this queue up to wait
    #: their turn. There is deliberately no per-call timeout -- see ToolGateway for
    #: why.
    mcp_max_concurrency: int = Field(default=1, alias="MCP_MAX_CONCURRENCY", ge=1, le=8)
    #: Hard cap on tool calls admitted at once (running + waiting for a worker).
    #: Once this many are already admitted, further calls are rejected immediately
    #: with a structured server_busy error instead of queuing indefinitely.
    mcp_max_queue_size: int = Field(default=20, alias="MCP_MAX_QUEUE_SIZE", ge=1, le=1000)

    log_level: Literal["DEBUG", "INFO", "WARNING", "ERROR"] = Field(
        default="INFO",
        alias="LOG_LEVEL",
    )
    log_file: Path | None = Field(default=None, alias="LOG_FILE")

    @field_validator(
        "exchange_email_address",
        "exchange_version",
        "exchange_impersonate_as",
        "attachment_root",
        "log_file",
        mode="before",
    )
    @classmethod
    def _blank_env_value_means_unset(cls, value: object) -> object:
        # .env.example ships optional settings as bare `NAME=` lines, and dotenv
        # delivers those as empty strings, not as missing keys. Without this,
        # Path-typed fields turn "" into Path(".") -- EXCHANGE_ATTACHMENT_ROOT
        # would silently sandbox attachments to the server's cwd instead of
        # staying disabled.
        if isinstance(value, str) and not value.strip():
            return None
        return value

    @model_validator(mode="after")
    def _reject_insecure_basic_auth(self) -> "Settings":
        if self.exchange_auth_type != "Basic" or self.exchange_allow_insecure_basic_auth:
            return self
        scheme = urlparse(self.exchange_server).scheme.lower()
        if scheme == "http":
            raise ValueError(
                "EXCHANGE_AUTH_TYPE=Basic sends credentials in the clear over "
                "EXCHANGE_SERVER=http://... . Use an https:// endpoint, or set "
                "EXCHANGE_ALLOW_INSECURE_BASIC_AUTH=true to explicitly opt into "
                "sending Basic auth over plain HTTP (e.g. for a local/test server)."
            )
        return self


@lru_cache(maxsize=1)
def get_settings() -> Settings:
    return Settings()  # type: ignore[call-arg]
