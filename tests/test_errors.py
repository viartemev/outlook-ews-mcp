from __future__ import annotations

from pydantic import BaseModel, ValidationError

from outlook_mcp.errors import APIError, normalize_exception, validation_error_from_pydantic


class SampleModel(BaseModel):
    value: int


def test_validation_error_from_pydantic() -> None:
    try:
        SampleModel(value="bad")
    except ValidationError as exc:
        error = validation_error_from_pydantic(exc)
    else:  # pragma: no cover
        raise AssertionError("validation error was expected")

    assert error.code == "validation_error"
    assert error.details[0]["field"] == "value"


def test_normalize_exception_preserves_api_error() -> None:
    error = APIError("custom", "boom")
    assert normalize_exception(error) is error


def test_normalize_exception_wraps_other_errors() -> None:
    error = normalize_exception(RuntimeError("boom, connect to https://exchange.internal/EWS failed"))
    assert error.code == "internal_error"
    assert "boom" not in error.message
    assert "exchange.internal" not in error.message
