from __future__ import annotations

import inspect
import json
import logging
import time
from collections.abc import Callable
from dataclasses import dataclass
from typing import Annotated, Any

import anyio
import anyio.to_thread
from mcp.types import CallToolResult, TextContent
from pydantic.fields import FieldInfo

from .config import Settings
from .errors import APIError

from .exchange_client import ExchangeClient
from .models import ExchangeModel

logger = logging.getLogger(__name__)


class ToolGateway:
    """Queues tool calls and keeps the blocking ones off the event loop.

    Three problems, one place to solve them.

    *Blocking.* FastMCP calls a synchronous tool function directly on the event
    loop thread, while our tool bodies do blocking HTTPS through exchangelib. A
    synchronous binding therefore freezes the whole server for the length of
    every Exchange round trip: no messages read, no finished responses written,
    no pings answered. Running the call in a worker thread keeps the transport
    alive.

    *Pile-up.* The MCP server dispatches every request concurrently, so a client
    that fires ten tool calls starts ten of them at once. One shared gate turns
    that into a bounded queue, and anyio hands waiters their turn in arrival
    order. MCP_MAX_QUEUE_SIZE bounds how many calls can be admitted at once
    (running plus waiting); once that many are already in, further calls are
    rejected immediately with a structured server_busy error instead of queuing
    forever.

    *Reads run in parallel, writes run alone.* Read-only tools (spec.read_only,
    the same flag published to clients as readOnlyHint) each take one of
    MCP_MAX_CONCURRENCY permits, so an agent asking for "the email, the folder
    list and the calendar" pays the slowest round trip, not the sum. A mutating
    tool first takes the writer lock (mutations never overlap each other), then
    collects every permit, so it also never overlaps a read: exchangelib's
    Protocol is thread-safe behind its session pool, but Account state (folder
    hierarchy, item caches) is not something this codebase audits for
    write-vs-read races -- exclusive writes make that audit unnecessary. A
    writer cannot be starved: readers pass through the writer lock as a
    turnstile first, so once a writer is waiting, later reads line up behind it
    instead of overtaking it at the semaphore.

    *Workers are never abandoned, and there is no per-call timeout.* Those are
    the same decision. Cancelling the wait on a worker does not stop it -- the
    thread keeps running and keeps the EWS session it checked out. exchangelib's
    session pool has a hard maximum and hands out sessions in a loop with no
    give-up path, so abandoned workers eventually hold every session and every
    later call blocks forever: a live process that answers nothing. The wait is
    therefore non-cancellable, and anyio.fail_after cannot fire across a
    non-cancellable wait -- rather than keep a timeout that silently never
    triggers, there is none.

    A call is instead bounded by EXCHANGE_TIMEOUT plus EXCHANGE_MAX_RETRY_WAIT_
    SECONDS. The account's retry_policy is FailFast (exchange_client/base.py), so
    every EWS service call raises on its first transient error instead of
    exchangelib retrying it internally with no cumulative time limit; ExchangeClient
    then retries read-only calls itself, bounded by a monotonic deadline built from
    EXCHANGE_MAX_RETRY_WAIT_SECONDS (see ExchangeClient._retry_read). Non-idempotent
    calls -- send, create, delete, and the like -- are never retried: an ambiguous
    failure could mean the operation already happened on the server. Overruns past
    the expected budget are logged.
    """

    def __init__(self, settings: Settings) -> None:
        self.max_workers = settings.mcp_max_concurrency
        self.max_queue_size = settings.mcp_max_queue_size
        self.expected_call_seconds = settings.exchange_timeout + settings.exchange_max_retry_wait
        # Safe to build outside a running event loop; anyio binds lazily.
        self._permits = anyio.Semaphore(self.max_workers)
        self._writer_turn = anyio.Lock()
        self._threads = anyio.CapacityLimiter(self.max_workers)
        self._admitted = 0

    async def run(
        self,
        name: str,
        call: Callable[[], tuple[Any, bool]],
        *,
        read_only: bool = False,
    ) -> tuple[Any, bool]:
        if self._admitted >= self.max_queue_size:
            logger.warning(
                "tool=%s rejected: %s calls already queued (MCP_MAX_QUEUE_SIZE=%s)",
                name,
                self._admitted,
                self.max_queue_size,
            )
            error = APIError(
                "server_busy",
                "too many tool calls are already queued; try again shortly",
                extra={"queued": self._admitted, "max_queue_size": self.max_queue_size},
            )
            return error.to_dict(), True
        self._admitted += 1
        try:
            if read_only:
                # Turnstile: while a writer holds the lock -- collecting permits
                # or running -- later readers line up here instead of slipping
                # past it to the semaphore. Without this, a steady stream of
                # reads starves the writer forever: it re-queues for each permit
                # and new readers keep overtaking it.
                async with self._writer_turn:
                    pass
                async with self._permits:
                    return await self._execute(name, call)
            async with self._writer_turn:
                acquired = 0
                try:
                    # Collect every permit so the write overlaps neither reads
                    # nor other writes. Readers already past the turnstile drain
                    # on their own; new ones wait at it.
                    for _ in range(self.max_workers):
                        await self._permits.acquire()
                        acquired += 1
                    return await self._execute(name, call)
                finally:
                    for _ in range(acquired):
                        self._permits.release()
        finally:
            self._admitted -= 1

    async def _execute(self, name: str, call: Callable[[], tuple[Any, bool]]) -> tuple[Any, bool]:
        started = time.monotonic()
        try:
            return await anyio.to_thread.run_sync(
                call, abandon_on_cancel=False, limiter=self._threads
            )
        finally:
            elapsed = time.monotonic() - started
            if elapsed > self.expected_call_seconds:
                logger.warning(
                    "tool=%s ran %.1fs, past the %.0fs Exchange budget; lower "
                    "EXCHANGE_TIMEOUT/EXCHANGE_MAX_RETRY_WAIT_SECONDS if calls "
                    "should give up sooner",
                    name,
                    elapsed,
                    self.expected_call_seconds,
                )


def _field_default(field: FieldInfo) -> Any:
    if field.is_required():
        return inspect.Parameter.empty
    return field.get_default(call_default_factory=True)


def _field_annotation(field: FieldInfo) -> Any:
    """The field's type, carrying its Field(...) constraints (ge/le/min_length/...).

    field.annotation alone drops that constraint metadata, so the JSON schema FastMCP
    publishes for the tool would advertise a bare type with no bounds even though the
    request model enforces them once the call comes in.
    """
    if field.metadata:
        return Annotated[(field.annotation, *field.metadata)]
    return field.annotation


@dataclass(frozen=True)
class ToolSpec:
    """Everything needed to register one Outlook MCP tool, in one place."""

    name: str
    description: str
    handler: Callable[[ExchangeClient, dict[str, Any]], Any]
    request_model: type[ExchangeModel] | None = None
    #: Validated by ToolRegistry for successful calls. FastMCP output-schema
    #: validation remains disabled because error payloads intentionally use a
    #: different shape and are returned with CallToolResult.isError=true.
    response_model: Any = None
    read_only: bool = False
    destructive: bool = False


def bind_mcp_tool(
    registry_call: Callable[[str, dict[str, Any]], tuple[Any, bool]],
    spec: ToolSpec,
    gateway: ToolGateway | None = None,
) -> Callable[..., Any]:
    """Build a FastMCP tool function with a schema derived from ``spec``.

    The returned function is a coroutine so FastMCP awaits it instead of running
    the blocking body on the event loop thread.
    """

    async def execute(**arguments: Any) -> CallToolResult:
        if gateway is None:
            payload, is_error = registry_call(spec.name, arguments)
        else:
            payload, is_error = await gateway.run(
                spec.name,
                lambda: registry_call(spec.name, arguments),
                read_only=spec.read_only,
            )
        structured = payload if isinstance(payload, dict) else {"result": payload}
        return CallToolResult(
            content=[TextContent(type="text", text=json.dumps(payload, ensure_ascii=False))],
            structuredContent=structured,
            isError=is_error,
        )

    execute.__name__ = spec.name
    execute.__doc__ = spec.description

    if spec.request_model is None:
        # FastMCP derives the tool schema from the function's signature/annotations,
        # so these must be patched onto the plain function object at runtime. The
        # return annotation is the bare CallToolResult type (not Annotated with a
        # payload model) so FastMCP skips output-schema validation entirely and lets
        # execute() report isError/structuredContent for both success and failure.
        execute.__signature__ = inspect.Signature(return_annotation=CallToolResult)  # type: ignore[attr-defined]
        execute.__annotations__ = {"return": CallToolResult}
        return execute

    parameters: list[inspect.Parameter] = []
    annotations: dict[str, Any] = {"return": CallToolResult}
    for field_name, field in spec.request_model.model_fields.items():
        annotation = _field_annotation(field)
        parameters.append(
            inspect.Parameter(
                field_name,
                inspect.Parameter.KEYWORD_ONLY,
                default=_field_default(field),
                annotation=annotation,
            )
        )
        annotations[field_name] = annotation

    execute.__signature__ = inspect.Signature(  # type: ignore[attr-defined]
        parameters, return_annotation=CallToolResult
    )
    execute.__annotations__ = annotations
    return execute


def register_mcp_tools(
    server: Any,
    registry: Any,
    tool_specs: list[ToolSpec],
    gateway: ToolGateway | None = None,
) -> None:
    """Register Outlook MCP tools on a FastMCP server instance.

    One gateway for the whole server, not one per tool: a per-tool queue would
    let N tools run N calls at once.
    """
    from mcp.types import ToolAnnotations

    for spec in tool_specs:
        tool_fn = bind_mcp_tool(registry.call, spec, gateway)
        annotations = ToolAnnotations(
            readOnlyHint=spec.read_only,
            destructiveHint=spec.destructive,
        )
        server.add_tool(
            tool_fn,
            name=spec.name,
            description=spec.description,
            annotations=annotations,
        )
