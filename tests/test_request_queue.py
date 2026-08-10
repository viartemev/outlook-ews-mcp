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


def _tool(registry: SlowRegistry, gateway: ToolGateway, *, read_only: bool = False):
    spec = ToolSpec(
        "get_email",
        "Get an email",
        lambda c, a: None,
        request_model=GetEmailRequest,
        read_only=read_only,
    )
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


def test_concurrent_mutating_calls_run_strictly_one_at_a_time() -> None:
    """Mutations are never safe to overlap: two writes can race on the same item,
    and Account state (folder hierarchy, item caches) is not audited for
    write-vs-read races. Even with permits to spare, writes go one by one."""
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


def test_read_only_calls_overlap_up_to_the_concurrency_setting() -> None:
    """Read-only tools each take one permit, so an agent asking for the email,
    the folder list and the calendar pays the slowest round trip, not the sum."""
    registry = SlowRegistry()
    tool_fn = _tool(registry, ToolGateway(_settings(MCP_MAX_CONCURRENCY=3)), read_only=True)

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


def test_calls_past_the_queue_cap_are_rejected_with_server_busy() -> None:
    """MCP_MAX_QUEUE_SIZE bounds admitted calls (running + waiting); once full,
    further calls must fail fast with a structured error instead of piling up
    behind an unbounded queue."""
    registry = SlowRegistry(delay=0.05)
    gateway = ToolGateway(_settings(MCP_MAX_CONCURRENCY=1, MCP_MAX_QUEUE_SIZE=2))
    tool_fn = _tool(registry, gateway)

    async def run_all() -> list:
        results = [None] * 5

        async def run_and_store(index: int) -> None:
            results[index] = await tool_fn(id=f"email-{index}")

        async with anyio.create_task_group() as tg:
            for index in range(5):
                tg.start_soon(run_and_store, index)
                await anyio.sleep(0)  # admit in submission order
        return results

    results = anyio.run(run_all)

    rejected = [r for r in results if r.isError and r.structuredContent["error"] == "server_busy"]
    accepted = [r for r in results if not r.isError]
    assert len(rejected) == 3
    assert len(accepted) == 2


def test_the_queue_cap_setting_has_a_default() -> None:
    gateway = ToolGateway(_settings())

    assert gateway.max_queue_size == 20


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


class KindedRegistry(SlowRegistry):
    """SlowRegistry that also tracks overlap per call kind (read vs write)."""

    def __init__(self, delay: float = 0.05) -> None:
        super().__init__(delay)
        self.active_kinds: list[str] = []
        self.observed_mixes: list[tuple[str, ...]] = []

    def call(self, name: str, arguments: dict) -> tuple[dict, bool]:
        kind = "read" if name.startswith("read") else "write"
        self.active_kinds.append(kind)
        self.observed_mixes.append(tuple(sorted(self.active_kinds)))
        try:
            return super().call(name, arguments)
        finally:
            self.active_kinds.remove(kind)


def _kinded_tools(registry: KindedRegistry, gateway: ToolGateway):
    read_spec = ToolSpec(
        "read_tool", "Read", lambda c, a: None, request_model=GetEmailRequest, read_only=True
    )
    write_spec = ToolSpec("write_tool", "Write", lambda c, a: None, request_model=GetEmailRequest)
    return (
        bind_mcp_tool(registry.call, read_spec, gateway),
        bind_mcp_tool(registry.call, write_spec, gateway),
    )


def test_a_mutating_call_never_overlaps_a_read() -> None:
    """The write path collects every permit before running, so at no observed
    moment may a write share the mailbox with any other call."""
    registry = KindedRegistry(delay=0.02)
    read_fn, write_fn = _kinded_tools(registry, ToolGateway(_settings(MCP_MAX_CONCURRENCY=4)))

    async def run_all() -> None:
        async with anyio.create_task_group() as tg:
            for index in range(4):
                tg.start_soon(lambda i=index: read_fn(id=f"email-r{i}"))
            tg.start_soon(lambda: write_fn(id="email-w"))
            for index in range(4, 8):
                tg.start_soon(lambda i=index: read_fn(id=f"email-r{i}"))

    anyio.run(run_all)

    for mix in registry.observed_mixes:
        if "write" in mix:
            assert mix == ("write",), f"a write ran alongside {mix}"
    assert any(mix == ("write",) for mix in registry.observed_mixes)


def test_a_writer_is_not_starved_by_a_stream_of_readers() -> None:
    """The failure mode the turnstile exists for: the writer re-queues for each
    permit, so without the turnstile every newly arriving reader overtakes it and
    a busy agent could postpone its one mutation indefinitely."""
    registry = SlowRegistry(delay=0.02)
    gateway = ToolGateway(_settings(MCP_MAX_CONCURRENCY=2))
    read_fn = _tool(registry, gateway, read_only=True)
    write_fn = _tool(registry, gateway)

    async def run_all() -> None:
        async with anyio.create_task_group() as tg:
            tg.start_soon(lambda: read_fn(id="read-0"))
            tg.start_soon(lambda: read_fn(id="read-1"))
            await anyio.sleep(0)
            tg.start_soon(lambda: write_fn(id="write"))
            await anyio.sleep(0)
            for index in range(2, 8):
                tg.start_soon(lambda i=index: read_fn(id=f"read-{i}"))
                await anyio.sleep(0)

    anyio.run(run_all)

    write_position = registry.started.index("write")
    assert write_position <= 3, (
        f"the write ran {write_position} calls in, after readers that arrived later: "
        f"{registry.started}"
    )


def test_reads_and_writes_share_the_admission_cap() -> None:
    """server_busy counts every admitted call -- running or waiting, read or
    write -- because the cap protects the process, not one of the two paths."""
    registry = SlowRegistry(delay=0.05)
    gateway = ToolGateway(_settings(MCP_MAX_CONCURRENCY=1, MCP_MAX_QUEUE_SIZE=2))
    read_fn = _tool(registry, gateway, read_only=True)
    write_fn = _tool(registry, gateway)

    async def run_all() -> list:
        results = [None] * 4

        async def run_and_store(index: int, fn) -> None:
            results[index] = await fn(id=f"email-{index}")

        async with anyio.create_task_group() as tg:
            tg.start_soon(run_and_store, 0, read_fn)
            await anyio.sleep(0)
            tg.start_soon(run_and_store, 1, write_fn)
            await anyio.sleep(0)
            tg.start_soon(run_and_store, 2, read_fn)
            await anyio.sleep(0)
            tg.start_soon(run_and_store, 3, write_fn)
        return results

    results = anyio.run(run_all)

    rejected = [r for r in results if r.isError and r.structuredContent["error"] == "server_busy"]
    assert len(rejected) == 2


def test_the_read_only_flag_agrees_with_the_facade_retry_classification() -> None:
    """spec.read_only now drives scheduling, not just the readOnlyHint shown to
    clients. ExchangeClient keeps its own independent idempotency classification:
    only calls that are safe to repeat go through _retry_read. A tool the gateway
    parallelizes as a read but the facade refuses to retry (or vice versa) means
    one of the two classifications is wrong."""
    import inspect as inspect_module

    from outlook_mcp.exchange_client import ExchangeClient
    from outlook_mcp.tool_specs import TOOL_SPECS

    # Read-only for the mailbox but not idempotent for the local machine: it
    # writes the download to disk, and a retry after a partial download would
    # leave an orphaned file (see the comment on ExchangeClient.get_attachment).
    retry_exempt = {"get_attachment"}

    for spec in TOOL_SPECS:
        if spec.name in retry_exempt:
            continue
        method = getattr(ExchangeClient, spec.handler.__name__)
        retries = "_retry_read" in inspect_module.getsource(method)
        assert retries == spec.read_only, (
            f"{spec.name}: read_only={spec.read_only} but the facade "
            f"{'uses' if retries else 'does not use'} _retry_read"
        )
