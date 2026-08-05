from __future__ import annotations

from ..models import (
    CreateEventRequest,
    DeleteEventRequest,
    FindFreeSlotsRequest,
    GetEventRequest,
    ListEventsRequest,
    RespondToInviteRequest,
    UpdateEventRequest,
)
from .common import tool_handler

list_events = tool_handler("list_events", ListEventsRequest)
get_event = tool_handler("get_event", GetEventRequest)
create_event = tool_handler("create_event", CreateEventRequest)
update_event = tool_handler("update_event", UpdateEventRequest)
delete_event = tool_handler("delete_event", DeleteEventRequest)
respond_to_invite = tool_handler("respond_to_invite", RespondToInviteRequest)
find_free_slots = tool_handler("find_free_slots", FindFreeSlotsRequest)
get_my_availability = tool_handler("get_my_availability", ListEventsRequest)
list_calendars = tool_handler("list_calendars")
