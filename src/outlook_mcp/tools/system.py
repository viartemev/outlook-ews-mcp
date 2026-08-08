from __future__ import annotations

from ..models import (
    CreateInboxRuleRequest,
    DeleteInboxRuleRequest,
    OutOfOfficeSettings,
    UpdateInboxRuleRequest,
)
from .common import tool_handler

ping_exchange = tool_handler("ping")
get_mailbox_info = tool_handler("get_mailbox_info")
get_out_of_office = tool_handler("get_out_of_office")
list_delegates = tool_handler("list_delegates")
list_inbox_rules = tool_handler("list_inbox_rules")
create_inbox_rule = tool_handler("create_inbox_rule", CreateInboxRuleRequest)
update_inbox_rule = tool_handler("update_inbox_rule", UpdateInboxRuleRequest)
delete_inbox_rule = tool_handler("delete_inbox_rule", DeleteInboxRuleRequest)
set_out_of_office = tool_handler("set_out_of_office", OutOfOfficeSettings)
