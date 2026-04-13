from __future__ import annotations

from typing import Any

from pydantic import ValidationError


class APIError(Exception):
    def __init__(
        self,
        code: str,
        message: str,
        *,
        details: list[dict[str, Any]] | None = None,
        extra: dict[str, Any] | None = None,
    ) -> None:
        super().__init__(message)
        self.code = code
        self.message = message
        self.details = details or []
        self.extra = extra or {}

    def to_dict(self) -> dict[str, Any]:
        payload: dict[str, Any] = {"error": self.code, "message": self.message}
        if self.details:
            payload["details"] = self.details
        payload.update(self.extra)
        return payload


class ExchangeUnavailableError(APIError):
    def __init__(self, message: str = "exchange is unavailable") -> None:
        super().__init__("exchange_unavailable", message)


class AuthFailedError(APIError):
    def __init__(self, message: str = "authentication failed") -> None:
        super().__init__("auth_failed", message)


class NotFoundError(APIError):
    def __init__(self, item_id: str) -> None:
        super().__init__("not_found", f"item {item_id} was not found", extra={"id": item_id})


class PermissionDeniedError(APIError):
    def __init__(self, message: str = "permission denied") -> None:
        super().__init__("permission_denied", message)


class TimeoutAPIError(APIError):
    def __init__(self, timeout_seconds: int) -> None:
        super().__init__(
            "timeout",
            f"exchange request timed out after {timeout_seconds} seconds",
            extra={"timeout_seconds": timeout_seconds},
        )


class ConflictError(APIError):
    def __init__(self, message: str = "item changed remotely") -> None:
        super().__init__("conflict", message)


def validation_error_from_pydantic(exc: ValidationError) -> APIError:
    details = []
    for error in exc.errors():
        details.append(
            {
                "field": ".".join(str(part) for part in error["loc"]),
                "reason": error["msg"],
            }
        )
    return APIError("validation_error", "invalid tool arguments", details=details)


def normalize_exception(exc: Exception) -> APIError:
    if isinstance(exc, APIError):
        return exc
    return APIError("internal_error", str(exc))
