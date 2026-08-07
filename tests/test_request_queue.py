from __future__ import annotations

import time

import anyio

from outlook_mcp.config import Settings
from outlook_mcp.mcp_tools import ToolGateway, ToolSpec, bind_mcp_tool
from outlook_mcp.models import GetEmailRequest


def _settings(**overrides) -> Settings:
    payload = {
        "EXCHANGE_SERVER": "https://mail.example.com/EWS/Exchange.asmx",
        "EXCHANGE_USERNAME": "user@example.com",
        "EXCHANGE_PASSWORD": "secret",
    }
    payload.update(overrides)
    return Settings(**payload)  # type: ignore[arg-type]


class SlowRegistry:
    """Stands in for ToolRegistry.call: blocking, and it records overlap."""

    def __init__(self, delay: float = 0.05) -> None:
        self.delay = delay
        self.active = 0
        self.max_active = 0
        self.started: list[str] = []

    def call(self, name: str, arguments: dict) -> tuple[dict, bool]:
        self.active += 1
        self.max_active = max(self.max_active, self.active)
        self.started.append(arguments.get("id", name))
        try:
            time.sleep(self.delay)  # blocking, exactly like an EWS round trip
            return {"id": arguments.get("id", name)}, False
        finally:
            self.active -= 1


def _tool(registry: SlowRegistry, gateway: ToolGateway):
    spec = ToolSpec("get_email", "Get an email", lambda c, a: None, request_model=GetEmailRequest)
    return bind_mcp_tool(registry.call, spec, gateway)


def test_a_running_tool_does_not_freeze_the_event_loop() -> None:
    """The whole point. FastMCP calls a synchronous tool function directly on the
    event loop thread, so a blocking binding stops the server from reading
    messages, writing finished responses or answering pings for the length of
    every Exchange round trip. With one, this counter stays at 0."""
    registry = SlowRegistry(delay=0.3)
    tool_fn = _tool(registry, ToolGateway(_settings()))
    ticks = 0

    async def heartbeat() -> None:
        nonlocal ticks
        while True:
            await anyio.sleep(0.01)
            ticks += 1

    async def scenario() -> None:
        async with anyio.create_task_group() as tg:
            tg.start_soon(heartbeat)
            await tool_fn(id="email-1")
            tg.cancel_scope.cancel()

    anyio.run(scenario)

    assert ticks > 5, f"event loop was blocked; it only ticked {ticks} times"


def test_concurrent_calls_run_strictly_one_at_a_time() -> None:
    registry = SlowRegistry()
    tool_fn = _tool(registry, ToolGateway(_settings()))

    async def run_all() -> None:
        async def one(index: int) -> None:
            await tool_fn(id=f"email-{index}")

        async with anyio.create_task_group() as tg:
            for index in range(6):
                tg.start_soon(one, index)

    anyio.run(run_all)

    assert registry.max_active == 1
    assert len(registry.started) == 6


def test_the_queue_is_first_in_first_out() -> None:
    registry = SlowRegistry()
    tool_fn = _tool(registry, ToolGateway(_settings()))

    async def run_all() -> None:
        async def one(index: int) -> None:
            await tool_fn(id=f"email-{index}")

        async with anyio.create_task_group() as tg:
            for index in range(6):
                tg.start_soon(one, index)
                await anyio.sleep(0)  # make submission order deterministic

    anyio.run(run_all)

    assert registry.started == [f"email-{index}" for index in range(6)]


def test_the_concurrency_setting_is_honoured() -> None:
    registry = SlowRegistry()
    tool_fn = _tool(registry, ToolGateway(_settings(MCP_MAX_CONCURRENCY=3)))

    async def run_all() -> None:
        async def one(index: int) -> None:
            await tool_fn(id=f"email-{index}")

        async with anyio.create_task_group() as tg:
            for index in range(9):
                tg.start_soon(one, index)

    anyio.run(run_all)

    assert registry.max_active == 3


def test_every_tool_shares_one_queue() -> None:
    """A per-tool semaphore would let N tools run N calls at once."""
    registry = SlowRegistry()
    gateway = ToolGateway(_settings())
    first = _tool(registry, gateway)
    second = bind_mcp_tool(
        registry.call,
        ToolSpec("get_thread", "Thread", lambda c, a: None, request_model=GetEmailRequest),
        gateway,
    )

    async def run_all() -> None:
        async with anyio.create_task_group() as tg:
            tg.start_soon(lambda: first(id="a"))
            tg.start_soon(lambda: second(id="b"))

    anyio.run(run_all)

    assert registry.max_active == 1


def test_slow_calls_do_not_leak_worker_threads() -> None:
    """A worker abandoned mid-call keeps running *and* keeps the EWS session it
    checked out. exchangelib's session pool has a hard maximum and hands out
    sessions in a loop with no give-up path, so abandoned workers eventually hold
    every session and every later call blocks forever -- a live process that
    answers nothing, curable only by restarting it."""
    import threading

    registry = SlowRegistry(delay=0.02)
    gateway = ToolGateway(_settings())
    tool_fn = _tool(registry, gateway)
    before = threading.active_count()

    async def run_all() -> None:
        for index in range(12):
            await tool_fn(id=f"email-{index}")

    anyio.run(run_all)

    grew_by = threading.active_count() - before
    assert grew_by <= gateway.max_workers


def test_an_overrunning_call_is_logged(caplog) -> None:
    registry = SlowRegistry(delay=0.05)
    gateway = ToolGateway(_settings())
    gateway.expected_call_seconds = 0.01
    tool_fn = _tool(registry, gateway)

    with caplog.at_level("WARNING"):
        anyio.run(lambda: tool_fn(id="email-1"))

    assert any("past the" in record.message for record in caplog.records)


def test_the_call_budget_comes_from_the_exchange_settings() -> None:
    gateway = ToolGateway(_settings(EXCHANGE_TIMEOUT=20, EXCHANGE_MAX_RETRY_WAIT_SECONDS=40))

    assert gateway.expected_call_seconds == 60


def test_every_registered_tool_is_awaitable(client, settings) -> None:
    """A tool registered as a plain function runs on the event loop thread and
    freezes the server for the length of the call."""
    from mcp.server.fastmcp.tools.base import _is_async_callable

    from outlook_mcp.server import build_mcp_server

    server = build_mcp_server(settings=settings, client=client)
    sync_tools = [
        name
        for name, tool in server._tool_manager._tools.items()
        if not _is_async_callable(tool.fn)
    ]

    assert sync_tools == []
