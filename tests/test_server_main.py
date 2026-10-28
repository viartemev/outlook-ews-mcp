from __future__ import annotations

import logging

import httpx
import pytest
from pydantic import SecretStr

from outlook_mcp import server as server_module
from outlook_mcp.config import Settings
from outlook_mcp.exchange_client import ExchangeClient
from outlook_mcp.server import BearerTokenMiddleware, build_mcp_server, main


def test_build_mcp_server_passes_http_host_and_port_to_constructor(
    client: ExchangeClient, settings: Settings
) -> None:
    """host/port belong on the FastMCP constructor, not on server.run() -
    passing them to run() raises TypeError against mcp==1.29.0's signature."""
    settings = settings.model_copy(update={"mcp_http_host": "0.0.0.0", "mcp_http_port": 9999})
    server = build_mcp_server(settings=settings, client=client)

    assert server.settings.host == "0.0.0.0"
    assert server.settings.port == 9999


def test_build_mcp_server_uses_distribution_name(client, settings) -> None:
    server = build_mcp_server(settings=settings, client=client)

    assert server.name == "outlook-ews-mcp"


def test_main_stdio_runs_with_no_arguments(
    monkeypatch, client: ExchangeClient, settings: Settings
) -> None:
    calls: list[tuple[tuple, dict]] = []

    class FakeServer:
        def run(self, *args, **kwargs):
            calls.append((args, kwargs))

    monkeypatch.setattr(server_module, "get_settings", lambda: settings)
    monkeypatch.setattr(server_module, "build_mcp_server", lambda settings: FakeServer())

    main()

    assert calls == [((), {})]


@pytest.mark.parametrize(
    ("transport", "app_factory"),
    [("streamable-http", "streamable_http_app"), ("sse", "sse_app")],
)
def test_main_serves_http_transports_behind_the_bearer_token(
    monkeypatch, client: ExchangeClient, settings: Settings, transport, app_factory
) -> None:
    """FastMCP.run() cannot require a token, so main() must serve the transport's
    ASGI app itself, wrapped in BearerTokenMiddleware."""
    settings = settings.model_copy(
        update={
            "mcp_transport": transport,
            "mcp_http_auth_token": SecretStr("s3cret"),
            "mcp_http_port": 9999,
        }
    )
    real_server = build_mcp_server(settings=settings, client=client)
    built: list[str] = []
    original = getattr(real_server, app_factory)
    monkeypatch.setattr(real_server, app_factory, lambda: built.append(app_factory) or original())
    runs: list[dict] = []
    monkeypatch.setattr(server_module, "get_settings", lambda: settings)
    monkeypatch.setattr(server_module, "build_mcp_server", lambda settings: real_server)
    monkeypatch.setattr(
        server_module.uvicorn, "run", lambda app, **kwargs: runs.append({"app": app, **kwargs})
    )

    main()

    assert built == [app_factory]
    assert len(runs) == 1
    assert isinstance(runs[0]["app"], BearerTokenMiddleware)
    assert (runs[0]["host"], runs[0]["port"]) == ("127.0.0.1", 9999)


def _token_client(app) -> httpx.AsyncClient:
    return httpx.AsyncClient(
        transport=httpx.ASGITransport(app=BearerTokenMiddleware(app, token="s3cret")),
        base_url="http://test",
    )


async def _ok_app(scope, receive, send):
    await send({"type": "http.response.start", "status": 200, "headers": []})
    await send({"type": "http.response.body", "body": b"ok"})


@pytest.mark.anyio
@pytest.mark.parametrize("header", [None, "Bearer wrong", "s3cret", "Basic s3cret"])
async def test_bearer_middleware_rejects_a_missing_or_wrong_token(header) -> None:
    headers = {"Authorization": header} if header else {}
    async with _token_client(_ok_app) as http:
        response = await http.post("/mcp", headers=headers)

    assert response.status_code == 401
    assert response.headers["www-authenticate"] == "Bearer"


@pytest.mark.anyio
async def test_bearer_middleware_passes_the_right_token_through() -> None:
    async with _token_client(_ok_app) as http:
        response = await http.post("/mcp", headers={"Authorization": "Bearer s3cret"})

    assert response.status_code == 200
    assert response.text == "ok"


@pytest.mark.anyio
async def test_real_streamable_http_app_refuses_an_unauthenticated_tool_call(
    client: ExchangeClient, settings: Settings
) -> None:
    app = build_mcp_server(settings=settings, client=client).streamable_http_app()
    call = {"jsonrpc": "2.0", "id": 1, "method": "tools/call", "params": {"name": "ping_exchange"}}

    async with _token_client(app) as http:
        response = await http.post("/mcp", json=call)

    assert response.status_code == 401


def test_configure_logging_does_not_leak_exchangelib_xml(settings: Settings, caplog) -> None:
    """exchangelib.util logs full SOAP request/response XML -- message bodies,
    recipients, base64 attachments -- at DEBUG, and re-embeds that same XML in
    a log.error() call on unexpected transport exceptions. Neither path should
    reach the logs, even with LOG_LEVEL=DEBUG."""
    settings = settings.model_copy(update={"log_level": "DEBUG"})
    server_module.configure_logging(settings)

    body_marker = "EMAIL-BODY-MARKER"
    xml_logger = logging.getLogger("exchangelib.util")
    xml_sub_logger = logging.getLogger("exchangelib.util.xml")

    with caplog.at_level(logging.DEBUG):
        xml_logger.debug("Request XML: %s", body_marker)
        xml_logger.error("TransportError: boom\nRequest XML: %s", body_marker)
        xml_sub_logger.debug("Request XML: %s", body_marker)

    assert body_marker not in caplog.text


def test_configure_logging_still_allows_outlook_mcp_debug_logs(settings: Settings, caplog) -> None:
    """Suppressing exchangelib's XML loggers must not silence our own logs --
    LOG_LEVEL=DEBUG should still surface outlook_mcp.* debug output."""
    settings = settings.model_copy(update={"log_level": "DEBUG"})
    server_module.configure_logging(settings)

    marker = "outlook-mcp-debug-marker"
    with caplog.at_level(logging.DEBUG):
        logging.getLogger("outlook_mcp.server").debug(marker)

    assert marker in caplog.text
