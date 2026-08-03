from __future__ import annotations

from datetime import datetime, time, timedelta
from typing import Any, Literal

from exchangelib import Attendee, CalendarItem
from exchangelib.recurrence import (
    WEEKDAY_NAMES,
    AbsoluteMonthlyPattern,
    AbsoluteYearlyPattern,
    DailyPattern,
    Recurrence,
    WeeklyPattern,
)

from ..models import (
    ActionResult,
    Attendee as ApiAttendee,
    AvailabilityResult,
    CalendarEvent,
    CalendarInfo,
    CreateEventRequest,
    CreateEventResult,
    DeleteEventRequest,
    FindFreeSlotsRequest,
    FreeSlot,
    GetEventRequest,
    ListEventsRequest,
    RecurrencePattern,
    RespondToInviteRequest,
    UpdateEventRequest,
    WorkHours,
)
from .base import BaseEWSBackend


class CalendarOperationsMixin(BaseEWSBackend):
    def _to_calendar_event(self, item: Any) -> CalendarEvent:
        attendees = [
            self._to_attendee(attendee)
            for attendee in (getattr(item, "required_attendees", None) or [])
            + (getattr(item, "optional_attendees", None) or [])
        ]
        organizer = self._email_address(getattr(item, "organizer", None))
        body_text, _ = self._extract_message_body(item)
        online_url = getattr(item, "meeting_workspace_url", None) or getattr(
            item, "net_show_url", None
        )
        recurrence = getattr(item, "recurrence", None)
        return CalendarEvent(
            id=item.id,
            subject=item.subject or "",
            start=item.start,
            end=item.end,
            location=getattr(item, "location", None),
            organizer=organizer,
            attendees=attendees,
            is_all_day=bool(getattr(item, "is_all_day", False)),
            is_recurring=bool(getattr(item, "is_recurring", False)),
            my_response=self._normalize_response(getattr(item, "my_response_type", None)),
            online_meeting_url=str(online_url) if online_url else None,
            body=body_text or None,
            reminder_minutes=getattr(item, "reminder_minutes_before_start", None),
            categories=list(getattr(item, "categories", None) or []),
            recurrence_pattern={"value": str(recurrence)} if recurrence else None,
            importance=self._normalize_importance(getattr(item, "importance", None)),
        )

    def _to_attendee(self, attendee: Any) -> ApiAttendee:
        mailbox = getattr(attendee, "mailbox", None) or attendee
        response = getattr(attendee, "response_type", None)
        address = self._email_address(mailbox)
        return ApiAttendee(
            email=address.email,
            name=address.name,
            response_type=self._normalize_response(response),
        )

    def _normalize_response(
        self, value: Any
    ) -> Literal["accept", "tentative", "decline", "unknown"]:
        normalized = str(value or "unknown").lower()
        mapping: dict[str, Literal["accept", "tentative", "decline", "unknown"]] = {
            "accept": "accept",
            "accepted": "accept",
            "organizer": "accept",
            "tentative": "tentative",
            "decline": "decline",
            "declined": "decline",
        }
        return mapping.get(normalized, "unknown")

    def _compute_free_slots(
        self, start: datetime, end: datetime, events: list[CalendarEvent]
    ) -> list[FreeSlot]:
        busy_intervals = sorted(
            (
                (max(event.start, start), min(event.end, end))
                for event in events
                # An appointment that ends when it starts occupies no time, and
                # one that ends before it starts says nothing we can trust. Both
                # exist in real mailboxes; letting either through would split the
                # free time around a moment nobody is actually busy.
                if event.end > event.start and event.end > start and event.start < end
            ),
            key=lambda interval: interval[0],
        )
        free_slots: list[FreeSlot] = []
        cursor = start
        for busy_start, busy_end in busy_intervals:
            if busy_start > cursor:
                free_slots.append(FreeSlot(start=cursor, end=busy_start, all_available=True))
            if busy_end > cursor:
                cursor = busy_end
        if cursor < end:
            free_slots.append(FreeSlot(start=cursor, end=end, all_available=True))
        return free_slots

    def list_events(self, request: ListEventsRequest) -> list[CalendarEvent]:
        folder = (
            self.account.calendar
            if not request.calendar_id
            else self._resolve_folder(request.calendar_id)
        )
        start = self._to_ews_datetime(request.start)
        end = self._to_ews_datetime(request.end)
        qs = folder.view(start=start, end=end)
        try:
            items = list(qs)
        except Exception as exc:  # noqa: BLE001
            raise self._map_exception(exc) from exc
        if not request.include_recurring:
            items = [item for item in items if not getattr(item, "is_recurring", False)]
        return [self._to_calendar_event(item) for item in items]

    def get_event(self, request: GetEventRequest) -> CalendarEvent:
        item = self._fetch_item(request.id, folder=self.account.calendar)
        return self._to_calendar_event(item)

    def _build_recurrence(self, pattern: RecurrencePattern, start: datetime) -> Recurrence:
        if pattern.type == "daily":
            ews_pattern = DailyPattern(interval=pattern.interval)
        elif pattern.type == "weekly":
            weekdays = (
                [day.capitalize() for day in pattern.days_of_week]
                if pattern.days_of_week
                else [WEEKDAY_NAMES[start.weekday()]]
            )
            ews_pattern = WeeklyPattern(interval=pattern.interval, weekdays=weekdays)
        elif pattern.type == "monthly":
            ews_pattern = AbsoluteMonthlyPattern(interval=pattern.interval, day_of_month=start.day)
        else:
            ews_pattern = AbsoluteYearlyPattern(day_of_month=start.day, month=start.month)

        boundary: dict[str, Any] = {"start": start.date()}
        if pattern.end_date is not None:
            boundary["end"] = pattern.end_date
        elif pattern.occurrences is not None:
            boundary["number"] = pattern.occurrences
        return Recurrence(pattern=ews_pattern, **boundary)

    def create_event(self, request: CreateEventRequest) -> CreateEventResult:
        folder = (
            self.account.calendar
            if not request.calendar_id
            else self._resolve_folder(request.calendar_id)
        )
        start = self._to_ews_datetime(request.start)
        end = self._to_ews_datetime(request.end)
        item = CalendarItem(
            account=self.account,
            folder=folder,
            subject=request.subject,
            start=start,
            end=end,
            location=request.location,
            body=request.body,
            required_attendees=[
                Attendee(mailbox=self._mailbox(address)) for address in request.attendees
            ],
            is_all_day=request.is_all_day,
            categories=request.categories,
            importance=request.importance.capitalize(),
            reminder_minutes_before_start=request.reminder_minutes,
            recurrence=self._build_recurrence(request.recurrence, start)
            if request.recurrence
            else None,
        )
        try:
            item.save(
                send_meeting_invitations="SendToAllAndSaveCopy"
                if request.attendees
                else "SendToNone"
            )
            return CreateEventResult(
                id=item.id or "",
                status="created",
                subject=request.subject,
                start=start,
                end=end,
                invite_sent=bool(request.attendees),
            )
        except Exception as exc:  # noqa: BLE001
            raise self._map_exception(exc) from exc

    def update_event(self, request: UpdateEventRequest) -> ActionResult:
        item = self._fetch_item(request.id, folder=self.account.calendar)
        updated_fields: list[str] = []
        save_fields: list[str] = []
        for field in ["subject", "start", "end", "location", "body", "reminder_minutes"]:
            value = getattr(request, field)
            if value is not None:
                if field in {"start", "end"}:
                    value = self._to_ews_datetime(value)
                target = "reminder_minutes_before_start" if field == "reminder_minutes" else field
                setattr(item, target, value)
                updated_fields.append(field)
                save_fields.append(target)
        if request.add_attendees:
            current = list(getattr(item, "required_attendees", None) or [])
            current.extend(
                Attendee(mailbox=self._mailbox(address)) for address in request.add_attendees
            )
            item.required_attendees = current
            updated_fields.append("add_attendees")
            if "required_attendees" not in save_fields:
                save_fields.append("required_attendees")
        if request.remove_attendees:
            remove_set = {address.lower() for address in request.remove_attendees}
            item.required_attendees = [
                attendee
                for attendee in getattr(item, "required_attendees", None) or []
                if getattr(getattr(attendee, "mailbox", None), "email_address", "").lower()
                not in remove_set
            ]
            updated_fields.append("remove_attendees")
            if "required_attendees" not in save_fields:
                save_fields.append("required_attendees")
        try:
            invitations = {
                "none": "SendToNone",
                "all": "SendToAllAndSaveCopy",
                "modified": "SendOnlyToChanged",
            }[request.send_updates]
            item.save(update_fields=save_fields or None, send_meeting_invitations=invitations)
            return ActionResult(id=request.id, status="updated", updated_fields=updated_fields)
        except Exception as exc:  # noqa: BLE001
            raise self._map_exception(exc, item_id=request.id) from exc

    def delete_event(self, request: DeleteEventRequest) -> ActionResult:
        item = self._fetch_item(request.id, folder=self.account.calendar)
        try:
            if request.notify_attendees and request.cancel_message:
                item.cancel(body=request.cancel_message)
            else:
                item.delete(
                    send_meeting_cancellations="SendToAllAndSaveCopy"
                    if request.notify_attendees
                    else "SendToNone"
                )
            return ActionResult(id=request.id, status="deleted")
        except Exception as exc:  # noqa: BLE001
            raise self._map_exception(exc, item_id=request.id) from exc

    def respond_to_invite(self, request: RespondToInviteRequest) -> ActionResult:
        item = self._fetch_item(request.id, folder=self.account.calendar)
        try:
            if request.response == "accept":
                item.accept(body=request.message)
            elif request.response == "tentative":
                item.tentatively_accept(body=request.message)
            else:
                item.decline(body=request.message)
            return ActionResult(id=request.id, status=request.response)
        except Exception as exc:  # noqa: BLE001
            raise self._map_exception(exc, item_id=request.id) from exc

    def find_free_slots(self, request: FindFreeSlotsRequest) -> list[FreeSlot]:
        start = self._to_ews_datetime(request.start)
        end = self._to_ews_datetime(request.end)
        try:
            views = list(
                self.account.protocol.get_free_busy_info(
                    accounts=[(str(address), "Required", False) for address in request.attendees],
                    start=start,
                    end=end,
                    merged_free_busy_interval=request.duration,
                    requested_view="DetailedMerged",
                )
            )
        except Exception as exc:  # noqa: BLE001
            raise self._map_exception(exc) from exc

        for view in views:
            if isinstance(view, Exception):
                raise self._map_exception(view)
        if not views:
            return []

        merged_per_attendee = [getattr(view, "merged", "") or "" for view in views]
        work_hours = self._parse_work_hours(request.work_hours)
        slots: list[FreeSlot] = []
        cursor = start
        interval = timedelta(minutes=request.duration)
        position = 0
        while cursor < end:
            slot_end = cursor + interval
            all_free = all(
                position < len(merged) and merged[position] == "0" for merged in merged_per_attendee
            )
            if all_free and self._within_work_hours(cursor, slot_end, work_hours):
                slots.append(
                    FreeSlot(start=cursor, end=slot_end, all_available=True, busy_attendees=[])
                )
            cursor = slot_end
            position += 1
        return slots

    def _parse_work_hours(self, work_hours: WorkHours | None) -> tuple[time, time] | None:
        if work_hours is None:
            return None
        return (
            datetime.strptime(work_hours.start, "%H:%M").time(),
            datetime.strptime(work_hours.end, "%H:%M").time(),
        )

    def _within_work_hours(
        self, slot_start: datetime, slot_end: datetime, work_hours: tuple[time, time] | None
    ) -> bool:
        if work_hours is None:
            return True
        day_start, day_end = work_hours
        return slot_start.time() >= day_start and slot_end.time() <= day_end

    def get_my_availability(self, request: ListEventsRequest) -> AvailabilityResult:
        events = self.list_events(request)
        busy_slots = [
            {"start": event.start, "end": event.end, "subject": event.subject} for event in events
        ]
        start = self._to_ews_datetime(request.start)
        end = self._to_ews_datetime(request.end)
        free_slots = self._compute_free_slots(start, end, events)
        return AvailabilityResult(free_slots=free_slots, busy_slots=busy_slots)

    def list_calendars(self) -> list[CalendarInfo]:
        return [
            CalendarInfo(
                id=self.account.calendar.id,
                name=self.account.calendar.name,
                is_default=True,
                owner_email=self.account.primary_smtp_address,
            )
        ]
