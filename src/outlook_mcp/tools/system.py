from __future__ import annotations

from ..models import OutOfOfficeSettings
from .common import tool_handler

ping_exchange = tool_handler("ping")
get_mailbox_info = tool_handler("get_mailbox_info")
get_out_of_office = tool_handler("get_out_of_office")
list_delegates = tool_handler("list_delegates")
set_out_of_office = tool_handler("set_out_of_office", OutOfOfficeSettings)
