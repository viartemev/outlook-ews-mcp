from __future__ import annotations

from outlook_mcp.tools.contacts import (
    create_contact,
    delete_contact,
    get_contact,
    search_contacts,
    update_contact,
)


def test_search_contacts(client) -> None:
    result = search_contacts(client, {"query": "ivan"})
    assert result[0]["display_name"] == "Ivan Ivanov"


def test_get_contact(client) -> None:
    result = get_contact(client, {"id": "contact-1"})
    assert result["id"] == "contact-1"
    assert result["email_addresses"][0]["address"] == "ivan@example.com"


def test_create_update_delete_contact(client) -> None:
    created = create_contact(client, {"display_name": "New Contact", "email": "new@example.com"})
    updated = update_contact(client, {"id": "contact-1", "display_name": "Updated"})
    deleted = delete_contact(client, {"id": "contact-1"})

    assert created["status"] == "created"
    assert updated["status"] == "updated"
    assert deleted["status"] == "deleted"
