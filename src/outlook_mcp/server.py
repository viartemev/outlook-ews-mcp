from __future__ import annotations

import logging
import sys
import time
from typing import Any

from .config import Settings, get_settings
from .errors import APIError, normalize_exception
from .exchange_client import ExchangeClient, build_default_backend
from .mcp_tools import normalize_tool_arguments, register_mcp_tools
from .tool_specs import TOOL_SPECS

logger = logging.getLogger(__name__)


def configure_logging(settings: Settings) -> None:
    handlers: list[logging.Handler]
    if (
        settings.log_file
        and str(settings.log_file).strip() not in {"", "."}
        and not settings.log_file.is_dir()
    ):
        handlers = [logging.FileHandler(settings.log_file)]
    else:
        handlers = [logging.StreamHandler(sys.stderr)]
    logging.basicConfig(
        level=getattr(logging, settings.log_level),
        format="%(asctime)s %(levelname)s %(name)s %(message)s",
        handlers=handlers,
    )


class ToolRegistry:
    def __init__(self, client: ExchangeClient) -> None:
        self.client = client
        self._specs = {spec.name: spec for spec in TOOL_SPECS}

    def call(
        self, name: str, arguments: dict[str, Any] | None = None
    ) -> tuple[dict[str, Any], bool]:
        arguments = normalize_tool_arguments(arguments)
        spec = self._specs.get(name)
        if not spec:
            error = APIError("not_found", f"unknown tool: {name}")
            return error.to_dict(), True

        started = time.perf_counter()
        try:
            result = spec.handler(self.client, arguments)
            duration_ms = round((time.perf_counter() - started) * 1000)
            logger.info("tool=%s status=ok duration_ms=%s", name, duration_ms)
            return result, False
        except Exception as exc:  # noqa: BLE001
            api_error = normalize_exception(exc)
            duration_ms = round((time.perf_counter() - started) * 1000)
            # Unclassified exceptions get a sanitized client-facing message (see
            # normalize_exception); log the real exception here, server-side only,
            # so it stays diagnosable without leaking internals to the MCP client.
            logger.warning(
                "tool=%s status=error duration_ms=%s error=%s",
                name,
                duration_ms,
                api_error.code,
                exc_info=exc if not isinstance(exc, APIError) else None,
            )
            return api_error.to_dict(), True


def build_registry(
    settings: Settings | None = None, client: ExchangeClient | None = None
) -> ToolRegistry:
    settings = settings or get_settings()
    configure_logging(settings)
    client = client or ExchangeClient(settings=settings, backend=build_default_backend(settings))
    return ToolRegistry(client)


def build_mcp_server(settings: Settings | None = None, client: ExchangeClient | None = None) -> Any:
    try:
        from mcp.server.fastmcp import FastMCP
    except ImportError as exc:  # pragma: no cover
        raise RuntimeError("mcp package is required to run the server") from exc

    registry = build_registry(settings=settings, client=client)
    server = FastMCP("outlook-mcp")
    register_mcp_tools(server, registry, TOOL_SPECS)
    return server


def main() -> None:
    settings = get_settings()
    server = build_mcp_server(settings=settings)
    transport = settings.mcp_transport

    if transport == "stdio":
        server.run()
        return

    if transport == "sse":  # pragma: no cover
        server.run(transport="sse", host=settings.mcp_sse_host, port=settings.mcp_sse_port)
        return

    raise RuntimeError(f"unsupported transport: {transport}")
