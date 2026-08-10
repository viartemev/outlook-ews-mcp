from __future__ import annotations

from ..models import (
    BulkDeleteEventsRequest,
    BulkRespondToInvitesRequest,
    CreateEventRequest,
    DeleteEventRequest,
    FindFreeSlotsRequest,
    GetEventRequest,
    ListEventsRequest,
    ListRoomsRequest,
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
delete_events = tool_handler("delete_events", BulkDeleteEventsRequest)
respond_to_invites = tool_handler("respond_to_invites", BulkRespondToInvitesRequest)
get_my_availability = tool_handler("get_my_availability", ListEventsRequest)
list_calendars = tool_handler("list_calendars")
list_room_lists = tool_handler("list_room_lists")
list_rooms = tool_handler("list_rooms", ListRoomsRequest)
