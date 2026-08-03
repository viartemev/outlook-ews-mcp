from __future__ import annotations

import logging

import uvicorn

from outlook_mcp import server as server_module
from outlook_mcp.config import Settings
from outlook_mcp.exchange_client import ExchangeClient
from outlook_mcp.server import build_mcp_server, main


def test_build_mcp_server_passes_sse_host_and_port_to_constructor(
    client: ExchangeClient, settings: Settings
) -> None:
    """host/port belong on the FastMCP constructor, not on server.run() -
    passing them to run() raises TypeError against mcp==1.29.0's signature."""
    settings = settings.model_copy(update={"mcp_sse_host": "0.0.0.0", "mcp_sse_port": 9999})
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


def test_main_sse_runs_via_uvicorn_with_bearer_middleware(
    monkeypatch, client: ExchangeClient, settings: Settings
) -> None:
    """main() must wrap the sse ASGI app in BearerTokenMiddleware and hand it to
    uvicorn.run directly - FastMCP.run(transport="sse") has no way to require a
    bearer token, so the sse branch bypasses it entirely."""
    settings = settings.model_copy(update={"mcp_transport": "sse", "mcp_sse_auth_token": "s3cret"})
    real_server = build_mcp_server(settings=settings, client=client)

    run_calls: list[dict] = []

    def fake_uvicorn_run(app, host=None, port=None):
        run_calls.append({"app": app, "host": host, "port": port})

    monkeypatch.setattr(server_module, "get_settings", lambda: settings)
    monkeypatch.setattr(server_module, "build_mcp_server", lambda settings: real_server)
    monkeypatch.setattr(uvicorn, "run", fake_uvicorn_run)

    main()

    assert len(run_calls) == 1
    call = run_calls[0]
    assert isinstance(call["app"], server_module.BearerTokenMiddleware)
    assert call["host"] == settings.mcp_sse_host
    assert call["port"] == settings.mcp_sse_port


def test_configure_logging_does_not_leak_exchangelib_xml(settings: Settings, caplog) -> None:
    """exchangelib.util logs full SOAP request/response XML -- message bodies,
    recipients, base64 attachments -- at DEBUG, and re-embeds that same XML in
    a log.error() call on unexpected transport exceptions. Neither path should
    reach the logs, even with LOG_LEVEL=DEBUG."""
    settings = settings.model_copy(update={"log_level": "DEBUG"})
    server_module.configure_logging(settings)

    secret = "SUPER-SECRET-EMAIL-BODY-MARKER"
    xml_logger = logging.getLogger("exchangelib.util")
    xml_sub_logger = logging.getLogger("exchangelib.util.xml")

    with caplog.at_level(logging.DEBUG):
        xml_logger.debug("Request XML: %s", secret)
        xml_logger.error("TransportError: boom\nRequest XML: %s", secret)
        xml_sub_logger.debug("Request XML: %s", secret)

    assert secret not in caplog.text


def test_configure_logging_still_allows_outlook_mcp_debug_logs(settings: Settings, caplog) -> None:
    """Suppressing exchangelib's XML loggers must not silence our own logs --
    LOG_LEVEL=DEBUG should still surface outlook_mcp.* debug output."""
    settings = settings.model_copy(update={"log_level": "DEBUG"})
    server_module.configure_logging(settings)

    marker = "outlook-mcp-debug-marker"
    with caplog.at_level(logging.DEBUG):
        logging.getLogger("outlook_mcp.server").debug(marker)

    assert marker in caplog.text
