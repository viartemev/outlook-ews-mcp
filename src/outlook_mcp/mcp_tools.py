from __future__ import annotations

import inspect
import json
from collections.abc import Callable
from typing import Any

from pydantic.fields import FieldInfo

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
    if field.default_factory is not None:
        return field.default_factory()
    return field.default


def bind_mcp_tool(
    registry_call: Callable[[str, dict[str, Any]], tuple[Any, bool]],
    name: str,
    description: str,
    request_model: type[ExchangeModel] | None = None,
) -> Callable[..., str]:
    """Build a FastMCP tool function with a schema derived from ``request_model``."""

    def execute(**arguments: Any) -> str:
        payload, is_error = registry_call(name, normalize_tool_arguments(arguments))
        if is_error:
            raise RuntimeError(json.dumps(payload, ensure_ascii=False))
        return json.dumps(payload, ensure_ascii=False)

    execute.__name__ = name
    execute.__doc__ = description

    if request_model is None:
        execute.__signature__ = inspect.Signature()
        execute.__annotations__ = {"return": str}
        return execute

    parameters: list[inspect.Parameter] = []
    annotations: dict[str, Any] = {"return": str}
    for field_name, field in request_model.model_fields.items():
        parameters.append(
            inspect.Parameter(
                field_name,
                inspect.Parameter.KEYWORD_ONLY,
                default=_field_default(field),
                annotation=field.annotation,
            )
        )
        annotations[field_name] = field.annotation

    execute.__signature__ = inspect.Signature(parameters, return_annotation=str)
    execute.__annotations__ = annotations
    return execute


def register_mcp_tools(server: Any, registry: Any, tool_specs: list[tuple[str, str, type[ExchangeModel] | None]]) -> None:
    """Register Outlook MCP tools on a FastMCP server instance."""
    for name, description, request_model in tool_specs:
        tool_fn = bind_mcp_tool(registry.call, name, description, request_model)
        server.add_tool(tool_fn, name=name, description=description)
