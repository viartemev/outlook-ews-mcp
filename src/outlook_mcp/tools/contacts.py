from __future__ import annotations

from pydantic import ValidationError

from ..errors import validation_error_from_pydantic
from ..exchange_client import ExchangeClient
from ..models import (
    CreateContactRequest,
    DeleteContactRequest,
    GetContactRequest,
    SearchContactsRequest,
    UpdateContactRequest,
    dump_model,
)


def search_contacts(client: ExchangeClient, arguments: dict) -> list[dict]:
    try:
        request = SearchContactsRequest.model_validate(arguments)
    except ValidationError as exc:
        raise validation_error_from_pydantic(exc) from exc
    return dump_model(client.search_contacts(request))


def get_contact(client: ExchangeClient, arguments: dict) -> dict:
    try:
        request = GetContactRequest.model_validate(arguments)
    except ValidationError as exc:
        raise validation_error_from_pydantic(exc) from exc
    return dump_model(client.get_contact(request))


def create_contact(client: ExchangeClient, arguments: dict) -> dict:
    try:
        request = CreateContactRequest.model_validate(arguments)
    except ValidationError as exc:
        raise validation_error_from_pydantic(exc) from exc
    return dump_model(client.create_contact(request))


def update_contact(client: ExchangeClient, arguments: dict) -> dict:
    try:
        request = UpdateContactRequest.model_validate(arguments)
    except ValidationError as exc:
        raise validation_error_from_pydantic(exc) from exc
    return dump_model(client.update_contact(request))


def delete_contact(client: ExchangeClient, arguments: dict) -> dict:
    try:
        request = DeleteContactRequest.model_validate(arguments)
    except ValidationError as exc:
        raise validation_error_from_pydantic(exc) from exc
    return dump_model(client.delete_contact(request))
