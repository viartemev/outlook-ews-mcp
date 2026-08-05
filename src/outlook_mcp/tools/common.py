from __future__ import annotations

from collections.abc import Callable
from typing import Any

from pydantic import ValidationError

from ..errors import validation_error_from_pydantic
from ..exchange_client import ExchangeClient
from ..models import ExchangeModel, dump_model

RequestHook = Callable[[ExchangeClient, ExchangeModel], None]


def tool_handler(
    method_name: str,
    request_model: type[ExchangeModel] | None = None,
    *,
    before: RequestHook | None = None,
) -> Callable[[ExchangeClient, dict[str, Any]], Any]:
    """Build a tool adapter with shared input validation and JSON serialization."""

    def handler(client: ExchangeClient, arguments: dict[str, Any]) -> Any:
        method = getattr(client, method_name)
        if request_model is None:
            return dump_model(method())
        try:
            request = request_model.model_validate(arguments)
        except ValidationError as exc:
            raise validation_error_from_pydantic(exc) from exc
        if before is not None:
            before(client, request)
        return dump_model(method(request))

    handler.__name__ = method_name
    return handler
