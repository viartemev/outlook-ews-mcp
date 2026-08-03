from __future__ import annotations

from pydantic import ValidationError

from ..errors import validation_error_from_pydantic
from ..exchange_client import ExchangeClient
from ..models import (
    CreateEventRequest,
    DeleteEventRequest,
    FindFreeSlotsRequest,
    GetEventRequest,
    ListEventsRequest,
    RespondToInviteRequest,
    UpdateEventRequest,
    dump_model,
)


def list_events(client: ExchangeClient, arguments: dict) -> list[dict]:
    try:
        request = ListEventsRequest.model_validate(arguments)
    except ValidationError as exc:
        raise validation_error_from_pydantic(exc) from exc
    return dump_model(client.list_events(request))


def get_event(client: ExchangeClient, arguments: dict) -> dict:
    try:
        request = GetEventRequest.model_validate(arguments)
    except ValidationError as exc:
        raise validation_error_from_pydantic(exc) from exc
    return dump_model(client.get_event(request))


def create_event(client: ExchangeClient, arguments: dict) -> dict:
    try:
        request = CreateEventRequest.model_validate(arguments)
    except ValidationError as exc:
        raise validation_error_from_pydantic(exc) from exc
    return dump_model(client.create_event(request))


def update_event(client: ExchangeClient, arguments: dict) -> dict:
    try:
        request = UpdateEventRequest.model_validate(arguments)
    except ValidationError as exc:
        raise validation_error_from_pydantic(exc) from exc
    return dump_model(client.update_event(request))


def delete_event(client: ExchangeClient, arguments: dict) -> dict:
    try:
        request = DeleteEventRequest.model_validate(arguments)
    except ValidationError as exc:
        raise validation_error_from_pydantic(exc) from exc
    return dump_model(client.delete_event(request))


def respond_to_invite(client: ExchangeClient, arguments: dict) -> dict:
    try:
        request = RespondToInviteRequest.model_validate(arguments)
    except ValidationError as exc:
        raise validation_error_from_pydantic(exc) from exc
    return dump_model(client.respond_to_invite(request))


def find_free_slots(client: ExchangeClient, arguments: dict) -> list[dict]:
    try:
        request = FindFreeSlotsRequest.model_validate(arguments)
    except ValidationError as exc:
        raise validation_error_from_pydantic(exc) from exc
    return dump_model(client.find_free_slots(request))


def get_my_availability(client: ExchangeClient, arguments: dict) -> dict:
    try:
        request = ListEventsRequest.model_validate(arguments)
    except ValidationError as exc:
        raise validation_error_from_pydantic(exc) from exc
    return dump_model(client.get_my_availability(request))


def list_calendars(client: ExchangeClient, arguments: dict) -> list[dict]:
    return dump_model(client.list_calendars())
