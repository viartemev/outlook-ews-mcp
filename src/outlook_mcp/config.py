from __future__ import annotations

from functools import lru_cache
from pathlib import Path
from typing import Literal
from urllib.parse import urlparse

from pydantic import Field, model_validator
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
    exchange_password: str = Field(alias="EXCHANGE_PASSWORD")
    exchange_email_address: str | None = Field(default=None, alias="EXCHANGE_EMAIL_ADDRESS")
    exchange_verify_ssl: bool = Field(default=True, alias="EXCHANGE_VERIFY_SSL")
    exchange_auth_type: Literal["NTLM", "Basic", "OAuth2"] = Field(
        default="NTLM",
        alias="EXCHANGE_AUTH_TYPE",
    )
    exchange_allow_insecure_basic_auth: bool = Field(
        default=False, alias="EXCHANGE_ALLOW_INSECURE_BASIC_AUTH"
    )
    exchange_version: str | None = Field(default=None, alias="EXCHANGE_VERSION")
    exchange_timeout: int = Field(default=30, alias="EXCHANGE_TIMEOUT", ge=1, le=300)
    exchange_max_retry_wait: int = Field(
        default=90, alias="EXCHANGE_MAX_RETRY_WAIT_SECONDS", ge=0, le=3600
    )
    exchange_timezone: str = Field(default="Europe/Moscow", alias="EXCHANGE_TIMEZONE")
    exchange_impersonate_as: str | None = Field(default=None, alias="EXCHANGE_IMPERSONATE_AS")
    attachment_max_size_mb: int = Field(default=10, alias="ATTACHMENT_MAX_SIZE_MB", ge=1, le=100)
    attachment_root: Path | None = Field(default=None, alias="EXCHANGE_ATTACHMENT_ROOT")

    mcp_transport: Literal["stdio", "sse"] = Field(default="stdio", alias="MCP_TRANSPORT")
    mcp_sse_host: str = Field(default="127.0.0.1", alias="MCP_SSE_HOST")
    mcp_sse_port: int = Field(default=8080, alias="MCP_SSE_PORT", ge=1, le=65535)

    log_level: Literal["DEBUG", "INFO", "WARNING", "ERROR"] = Field(
        default="INFO",
        alias="LOG_LEVEL",
    )
    log_file: Path | None = Field(default=None, alias="LOG_FILE")

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
