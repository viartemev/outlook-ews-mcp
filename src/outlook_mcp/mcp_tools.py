from __future__ import annotations

import inspect
import json
from collections.abc import Callable
from dataclasses import dataclass
from typing import Annotated, Any

from mcp.types import CallToolResult, TextContent
from pydantic.fields import FieldInfo

from .exchange_client import ExchangeClient
from .models import ExchangeModel


def normalize_tool_arguments(arguments: dict[str, Any] | None) -> dict[str, Any]:
    """Unwrap legacy clients that nest tool args under a single ``kwargs`` key."""
    arguments = dict(arguments or {})
    if len(arguments) == 1 and "kwargs" in arguments and isinstance(arguments["kwargs"], dict):
        return dict(arguments["kwargs"])
    return arguments


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
    #: Documents the handler's success payload shape; not enforced as an MCP output
    #: schema (see bind_mcp_tool — every tool reports isError via CallToolResult, and
    #: error payloads don't share this shape, so FastMCP's schema validation can't apply).
    response_model: Any = None
    read_only: bool = False
    destructive: bool = False


def bind_mcp_tool(
    registry_call: Callable[[str, dict[str, Any]], tuple[Any, bool]],
    spec: ToolSpec,
) -> Callable[..., Any]:
    """Build a FastMCP tool function with a schema derived from ``spec``."""

    def execute(**arguments: Any) -> CallToolResult:
        payload, is_error = registry_call(spec.name, normalize_tool_arguments(arguments))
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


def register_mcp_tools(server: Any, registry: Any, tool_specs: list[ToolSpec]) -> None:
    """Register Outlook MCP tools on a FastMCP server instance."""
    from mcp.types import ToolAnnotations

    for spec in tool_specs:
        tool_fn = bind_mcp_tool(registry.call, spec)
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
