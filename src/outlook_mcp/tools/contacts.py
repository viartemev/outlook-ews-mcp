from __future__ import annotations

from ..models import (
    CreateContactRequest,
    DeleteContactRequest,
    GetContactRequest,
    SearchContactsRequest,
    UpdateContactRequest,
)
from .common import tool_handler

search_contacts = tool_handler("search_contacts", SearchContactsRequest)
get_contact = tool_handler("get_contact", GetContactRequest)
create_contact = tool_handler("create_contact", CreateContactRequest)
update_contact = tool_handler("update_contact", UpdateContactRequest)
delete_contact = tool_handler("delete_contact", DeleteContactRequest)
