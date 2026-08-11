from __future__ import annotations

import logging
import secrets
import sys
import time
from typing import Any

from pydantic import TypeAdapter

from . import telemetry
from .config import Settings, get_settings
from .errors import APIError, normalize_exception
from .exchange_client import ExchangeClient, build_default_backend
from .mcp_tools import ToolGateway, register_mcp_tools
from .models import dump_model
from .tool_specs import TOOL_SPECS

logger = logging.getLogger(__name__)


class BearerTokenMiddleware:
    """Raw ASGI middleware gating every request behind a static bearer token.

    Deliberately not Starlette's BaseHTTPMiddleware: that buffers the whole
    response before sending it, which breaks the long-lived text/event-stream
    connection SSE depends on. This passes scope/receive/send straight through
    once authorized, so streaming is untouched.
    """

    def __init__(self, app: Any, token: str) -> None:
        self._app = app
        self._expected = f"Bearer {token}"

    async def __call__(self, scope: dict[str, Any], receive: Any, send: Any) -> None:
        if scope["type"] != "http":
            await self._app(scope, receive, send)
            return

        headers = dict(scope.get("headers") or [])
        presented = headers.get(b"authorization", b"").decode("latin-1")
        if not secrets.compare_digest(presented, self._expected):
            from starlette.responses import PlainTextResponse

            response = PlainTextResponse(
                "Unauthorized", status_code=401, headers={"WWW-Authenticate": "Bearer"}
            )
            await response(scope, receive, send)
            return

        await self._app(scope, receive, send)


def configure_logging(settings: Settings) -> None:
    handlers: list[logging.Handler]
    if settings.log_file is not None and not settings.log_file.is_dir():
        handlers = [logging.FileHandler(settings.log_file)]
    else:
        handlers = [logging.StreamHandler(sys.stderr)]
    # LOG_LEVEL only applies to our own loggers. The root level stays at the
    # logging-module default (WARNING) so third-party libraries -- notably
    # exchangelib -- don't get promoted to DEBUG just because an operator
    # wants verbose outlook_mcp logs.
    logging.basicConfig(
        format="%(asctime)s %(levelname)s %(name)s %(message)s",
        handlers=handlers,
    )
    logging.getLogger("outlook_mcp").setLevel(getattr(logging, settings.log_level))
    # exchangelib.util logs full SOAP request/response XML -- message bodies,
    # recipients, base64 attachment contents -- at DEBUG, and also embeds that
    # same XML in a log.error() call whenever an unexpected transport exception
    # is raised. That error call fires regardless of the configured level, so
    # these loggers are always pinned to CRITICAL rather than left to inherit
    # whatever level the operator picks for LOG_LEVEL.
    logging.getLogger("exchangelib.util").setLevel(logging.CRITICAL)
    logging.getLogger("exchangelib.util.xml").setLevel(logging.CRITICAL)


class ToolRegistry:
    def __init__(self, client: ExchangeClient) -> None:
        self.client = client
        self._specs = {spec.name: spec for spec in TOOL_SPECS}
        self._response_adapters = {
            spec.name: TypeAdapter(spec.response_model)
            for spec in TOOL_SPECS
            if spec.response_model is not None
        }

    def call(self, name: str, arguments: dict[str, Any] | None = None) -> tuple[Any, bool]:
        arguments = dict(arguments or {})
        spec = self._specs.get(name)
        if not spec:
            error = APIError("not_found", f"unknown tool: {name}")
            return error.to_dict(), True

        started = time.perf_counter()
        # The hook in exchange_client/base.py reports each EWS HTTP request into
        # this accumulator; duration_ms minus ews_ms is the time spent in this
        # codebase rather than on the wire.
        ews = telemetry.start_collecting()
        try:
            result = spec.handler(self.client, arguments)
            adapter = self._response_adapters.get(name)
            if adapter is not None:
                result = dump_model(adapter.validate_python(result))
            duration_ms = round((time.perf_counter() - started) * 1000)
            logger.info(
                "tool=%s status=ok duration_ms=%s ews_requests=%s ews_ms=%s",
                name,
                duration_ms,
                ews.requests,
                round(ews.total_ms),
            )
            return result, False
        except Exception as exc:  # noqa: BLE001
            api_error = normalize_exception(exc)
            duration_ms = round((time.perf_counter() - started) * 1000)
            logger.warning(
                "tool=%s status=error duration_ms=%s ews_requests=%s ews_ms=%s error=%s "
                "exception_type=%s",
                name,
                duration_ms,
                ews.requests,
                round(ews.total_ms),
                api_error.code,
                type(exc).__name__,
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

    settings = settings or get_settings()
    registry = build_registry(settings=settings, client=client)
    server = FastMCP("outlook-ews-mcp", host=settings.mcp_sse_host, port=settings.mcp_sse_port)
    register_mcp_tools(server, registry, TOOL_SPECS, ToolGateway(settings))
    return server


def main() -> None:
    settings = get_settings()
    server = build_mcp_server(settings=settings)
    transport = settings.mcp_transport

    if transport == "stdio":
        server.run()
        return

    if transport == "sse":  # pragma: no cover
        import uvicorn

        assert settings.mcp_sse_auth_token  # enforced by Settings validation
        app = BearerTokenMiddleware(server.sse_app(), token=settings.mcp_sse_auth_token)
        uvicorn.run(app, host=settings.mcp_sse_host, port=settings.mcp_sse_port)
        return

    raise RuntimeError(f"unsupported transport: {transport}")
