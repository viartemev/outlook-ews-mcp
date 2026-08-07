from __future__ import annotations

import logging

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


def test_main_sse_runs_with_transport_only(
    monkeypatch, client: ExchangeClient, settings: Settings
) -> None:
    """Regression test: main() must call server.run(transport="sse") without
    host/port kwargs - mcp==1.29.0's FastMCP.run() only accepts transport and
    mount_path, and would raise TypeError if called the old way."""
    settings = settings.model_copy(update={"mcp_transport": "sse"})
    calls: list[tuple[tuple, dict]] = []

    class FakeServer:
        def run(self, *args, **kwargs):
            calls.append((args, kwargs))

    monkeypatch.setattr(server_module, "get_settings", lambda: settings)
    monkeypatch.setattr(server_module, "build_mcp_server", lambda settings: FakeServer())

    main()

    assert calls == [((), {"transport": "sse"})]


def test_main_sse_transport_does_not_raise_with_real_fastmcp(
    monkeypatch, client: ExchangeClient, settings: Settings
) -> None:
    """End-to-end check with the real FastMCP.run() signature: construct the
    real server, monkeypatch just the network-binding uvicorn call, and
    confirm main() reaches it without a TypeError from a bad kwarg."""
    settings = settings.model_copy(update={"mcp_transport": "sse"})
    real_server = build_mcp_server(settings=settings, client=client)

    run_calls: list[dict] = []

    async def fake_run_sse_async(self, mount_path=None):
        run_calls.append({"mount_path": mount_path})

    monkeypatch.setattr(server_module, "get_settings", lambda: settings)
    monkeypatch.setattr(server_module, "build_mcp_server", lambda settings: real_server)
    monkeypatch.setattr(type(real_server), "run_sse_async", fake_run_sse_async)

    main()

    assert run_calls == [{"mount_path": None}]


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
