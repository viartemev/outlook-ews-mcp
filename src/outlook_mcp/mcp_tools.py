from __future__ import annotations

import inspect
import json
from collections.abc import Callable
from dataclasses import dataclass
from typing import Any

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


@dataclass(frozen=True)
class ToolSpec:
    """Everything needed to register one Outlook MCP tool, in one place."""

    name: str
    description: str
    handler: Callable[[ExchangeClient, dict[str, Any]], Any]
    request_model: type[ExchangeModel] | None = None
    response_model: Any = None
    read_only: bool = False
    destructive: bool = False


def bind_mcp_tool(
    registry_call: Callable[[str, dict[str, Any]], tuple[Any, bool]],
    spec: ToolSpec,
) -> Callable[..., Any]:
    """Build a FastMCP tool function with a schema derived from ``spec``."""

    def execute(**arguments: Any) -> Any:
        payload, is_error = registry_call(spec.name, normalize_tool_arguments(arguments))
        if is_error:
            raise RuntimeError(json.dumps(payload, ensure_ascii=False))
        return payload

    execute.__name__ = spec.name
    execute.__doc__ = spec.description

    return_annotation = spec.response_model if spec.response_model is not None else Any

    if spec.request_model is None:
        # FastMCP derives the tool schema from the function's signature/annotations,
        # so these must be patched onto the plain function object at runtime.
        execute.__signature__ = inspect.Signature(return_annotation=return_annotation)  # type: ignore[attr-defined]
        execute.__annotations__ = {"return": return_annotation}
        return execute

    parameters: list[inspect.Parameter] = []
    annotations: dict[str, Any] = {"return": return_annotation}
    for field_name, field in spec.request_model.model_fields.items():
        parameters.append(
            inspect.Parameter(
                field_name,
                inspect.Parameter.KEYWORD_ONLY,
                default=_field_default(field),
                annotation=field.annotation,
            )
        )
        annotations[field_name] = field.annotation

    execute.__signature__ = inspect.Signature(  # type: ignore[attr-defined]
        parameters, return_annotation=return_annotation
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
