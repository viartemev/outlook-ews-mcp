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

from ..errors import APIError
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
            reminder_minutes=(
                getattr(item, "reminder_minutes_before_start", None)
                if getattr(item, "reminder_is_set", False)
                else None
            ),
            categories=list(getattr(item, "categories", None) or []),
            recurrence_pattern={"value": str(recurrence)} if recurrence else None,
            importance=self._normalize_importance(getattr(item, "importance", None)),
            free_busy_status=self._normalize_free_busy_status(
                getattr(item, "legacy_free_busy_status", None)
            ),
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

    def _calendar_folder(self, calendar_id: str | None) -> Any:
        folder = self.account.calendar if not calendar_id else self._resolve_folder(calendar_id)
        folder_class = getattr(folder, "folder_class", None)
        if calendar_id and folder_class is not None and folder_class != self._CALENDAR_FOLDER_CLASS:
            raise APIError(
                "validation_error",
                "calendar_id does not reference a calendar folder",
                details=[{"field": "calendar_id", "reason": "folder is not IPF.Appointment"}],
            )
        return folder

    def list_events(self, request: ListEventsRequest) -> list[CalendarEvent]:
        folder = self._calendar_folder(request.calendar_id)
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
        item = self._fetch_item(
            request.id,
            folder=self._calendar_folder(request.calendar_id),
            expected_type=CalendarItem,
        )
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

    def _all_day_bounds(self, start: Any, end: Any) -> tuple[Any, Any]:
        """Floor start/end to midnight in Exchange's own timezone.

        EWS all-day appointments use an exclusive end date (a single day off
        Aug 3rd is start=Aug3T00:00, end=Aug4T00:00). Flooring the caller's end
        to its own date's midnight -- rather than to 23:59:59 on that date --
        keeps that convention: a same-day start/end becomes a one-day event,
        and a start/end already a day apart stays a one-day event instead of
        silently growing by almost a full extra day.
        """
        start_date = start.date()
        end_date = end.date() if end.date() > start_date else start_date + timedelta(days=1)
        tz = self.account.default_timezone
        start_bound = self._to_ews_datetime(datetime.combine(start_date, time.min, tzinfo=tz))
        end_bound = self._to_ews_datetime(datetime.combine(end_date, time.min, tzinfo=tz))
        return start_bound, end_bound

    def create_event(self, request: CreateEventRequest) -> CreateEventResult:
        folder = self._calendar_folder(request.calendar_id)
        start = self._to_ews_datetime(request.start)
        end = self._to_ews_datetime(request.end)
        if request.is_all_day:
            start, end = self._all_day_bounds(start, end)
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
            reminder_is_set=request.reminder_minutes is not None,
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
        item = self._fetch_item(
            request.id,
            folder=self._calendar_folder(request.calendar_id),
            expected_type=CalendarItem,
        )
        resolved_start = self._to_ews_datetime(request.start) if request.start else None
        resolved_end = self._to_ews_datetime(request.end) if request.end else None
        if getattr(item, "is_all_day", False):
            # All-day appointments are midnight-to-midnight with an exclusive end
            # date; floor any caller-supplied bound to midnight of its own day so a
            # partial update can't drift the event to start/end mid-day.
            tz = self.account.default_timezone
            if resolved_start is not None:
                resolved_start = self._to_ews_datetime(
                    datetime.combine(resolved_start.date(), time.min, tzinfo=tz)
                )
            if resolved_end is not None:
                resolved_end = self._to_ews_datetime(
                    datetime.combine(resolved_end.date(), time.min, tzinfo=tz)
                )
        if resolved_start is not None or resolved_end is not None:
            # Only touch the existing item's start/end when one side of the range
            # is actually changing -- a lone new start/end must still be checked
            # against the counterpart already on the event, not just against itself.
            effective_start = resolved_start if resolved_start is not None else item.start
            effective_end = resolved_end if resolved_end is not None else item.end
            if effective_end <= effective_start:
                raise APIError(
                    "validation_error",
                    "end must be greater than start",
                    details=[
                        {
                            "field": "start/end",
                            "reason": "resulting event range is empty or reversed",
                        }
                    ],
                )
        updated_fields: list[str] = []
        save_fields: list[str] = []
        for field in ["subject", "start", "end"]:
            value = getattr(request, field)
            if value is not None:
                if field == "start":
                    value = resolved_start
                elif field == "end":
                    value = resolved_end
                setattr(item, field, value)
                updated_fields.append(field)
                save_fields.append(field)
        # location/body/reminder_minutes are nullable clear targets: a field present
        # in the request (even as null) means "clear it"; an omitted field means
        # "leave it untouched". Mirrors the fields_set handling in update_contact.
        fields_set = request.model_fields_set
        for request_field, item_field in [("location", "location"), ("body", "body")]:
            if request_field not in fields_set:
                continue
            setattr(item, item_field, getattr(request, request_field))
            updated_fields.append(request_field)
            save_fields.append(item_field)
        if "reminder_minutes" in fields_set:
            reminder_minutes = request.reminder_minutes
            item.reminder_is_set = reminder_minutes is not None
            item.reminder_minutes_before_start = reminder_minutes or 0
            updated_fields.append("reminder_minutes")
            save_fields.extend(["reminder_is_set", "reminder_minutes_before_start"])
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
        if not save_fields:
            # Nothing actually changed -- saving anyway would still fire attendee
            # notifications per request.send_updates (default "all") for a no-op update.
            return ActionResult(id=request.id, status="updated", updated_fields=updated_fields)
        try:
            invitations = {
                "none": "SendToNone",
                "all": "SendToAllAndSaveCopy",
                "modified": "SendOnlyToChanged",
            }[request.send_updates]
            item.save(update_fields=save_fields, send_meeting_invitations=invitations)
            return ActionResult(id=request.id, status="updated", updated_fields=updated_fields)
        except Exception as exc:  # noqa: BLE001
            raise self._map_exception(exc, item_id=request.id) from exc

    def delete_event(self, request: DeleteEventRequest) -> ActionResult:
        item = self._fetch_item(
            request.id,
            folder=self._calendar_folder(request.calendar_id),
            expected_type=CalendarItem,
        )
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
        item = self._fetch_item(
            request.id,
            folder=self._calendar_folder(request.calendar_id),
            expected_type=CalendarItem,
        )
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
        # EWS's free/busy sampling granularity is bounded to [5, 1440] minutes and
        # is independent of the requested meeting duration -- a 1-minute meeting
        # still has to be checked against 5-minute-wide free/busy samples.
        sample_minutes = max(request.duration, 5)
        try:
            views = list(
                self.account.protocol.get_free_busy_info(
                    accounts=[(str(address), "Required", False) for address in request.attendees],
                    start=start,
                    end=end,
                    merged_free_busy_interval=sample_minutes,
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
        sample_interval = timedelta(minutes=sample_minutes)
        # Stop once a full-length slot would no longer fit before `end` -- a
        # truncated final slot would be shorter than the requested duration and
        # so isn't actually usable for scheduling a meeting of that length.
        while cursor + interval <= end:
            slot_end = cursor + interval
            position = (cursor - start) // sample_interval
            all_free = all(
                position < len(merged) and merged[position] == "0" for merged in merged_per_attendee
            )
            if all_free and self._within_work_hours(cursor, slot_end, work_hours):
                slots.append(
                    FreeSlot(start=cursor, end=slot_end, all_available=True, busy_attendees=[])
                )
            cursor = slot_end
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
        # "Free" appointments are placeholders (e.g. reminders, holidays) that
        # don't actually block time -- only exclude those from busy time, the
        # same threshold find_free_slots applies to the EWS merged free/busy view.
        busy_events = [event for event in events if event.free_busy_status != "free"]
        busy_slots = [
            {"start": event.start, "end": event.end, "subject": event.subject}
            for event in busy_events
        ]
        start = self._to_ews_datetime(request.start)
        end = self._to_ews_datetime(request.end)
        free_slots = self._compute_free_slots(start, end, busy_events)
        return AvailabilityResult(free_slots=free_slots, busy_slots=busy_slots)

    #: Standard EWS folder class for calendar folders.
    _CALENDAR_FOLDER_CLASS = "IPF.Appointment"

    def list_calendars(self) -> list[CalendarInfo]:
        try:
            default_id = self.account.calendar.id
            # Unlike _get_folder_by_id (which avoids walk() because it already knows the
            # target id), discovering *every* calendar folder in the tree has no id to look
            # up -- walk() is the only way. It costs one batched Deep-traversal FindFolder
            # call, cached on `self.account.root` for the process lifetime, not one call per
            # folder, so it doesn't carry the same scaling cost as issue #5 was about.
            folders = [
                folder
                for folder in self.account.root.walk()
                if getattr(folder, "folder_class", None) == self._CALENDAR_FOLDER_CLASS
            ]
        except Exception as exc:  # noqa: BLE001
            raise self._map_exception(exc) from exc
        return [
            CalendarInfo(
                id=folder.id,
                name=folder.name,
                is_default=folder.id == default_id,
                owner_email=self.account.primary_smtp_address,
            )
            for folder in folders
        ]
