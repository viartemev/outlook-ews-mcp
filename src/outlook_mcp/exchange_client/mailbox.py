from __future__ import annotations

from typing import Literal, cast

from exchangelib.settings import OofSettings

from ..models import DelegateInfo, DelegatePermissionLevels, OutOfOfficeSettings
from .base import BaseEWSBackend

#: EWS uses capitalised states; the API uses lowercase like every other enum here.
_STATE_TO_EWS = {"disabled": "Disabled", "enabled": "Enabled", "scheduled": "Scheduled"}
_STATE_FROM_EWS = {value: key for key, value in _STATE_TO_EWS.items()}
_AUDIENCE_TO_EWS = {"none": "None", "known": "Known", "all": "All"}
_AUDIENCE_FROM_EWS = {value: key for key, value in _AUDIENCE_TO_EWS.items()}


class MailboxSettingsMixin(BaseEWSBackend):
    def list_delegates(self) -> list[DelegateInfo]:
        try:
            delegates = self.account.delegates
        except Exception as exc:  # noqa: BLE001
            raise self._map_exception(exc) from exc
        result: list[DelegateInfo] = []
        for delegate in delegates:
            user = getattr(delegate, "user_id", None)
            permissions = getattr(delegate, "delegate_permissions", None)

            def _level(field: str) -> str:
                return str(getattr(permissions, field, None) or "None")

            result.append(
                DelegateInfo(
                    email=getattr(user, "primary_smtp_address", None),
                    display_name=getattr(user, "display_name", None),
                    permissions=DelegatePermissionLevels(
                        calendar=_level("calendar_folder_permission_level"),
                        inbox=_level("inbox_folder_permission_level"),
                        tasks=_level("tasks_folder_permission_level"),
                        contacts=_level("contacts_folder_permission_level"),
                    ),
                    receives_copies_of_meeting_messages=bool(
                        getattr(delegate, "receive_copies_of_meeting_messages", False)
                    ),
                    can_view_private_items=bool(getattr(delegate, "view_private_items", False)),
                )
            )
        return result

    def get_out_of_office(self) -> OutOfOfficeSettings:
        try:
            oof = self.account.oof_settings
        except Exception as exc:  # noqa: BLE001
            raise self._map_exception(exc) from exc
        state = _STATE_FROM_EWS.get(getattr(oof, "state", "Disabled"), "disabled")
        audience = _AUDIENCE_FROM_EWS.get(getattr(oof, "external_audience", "All"), "all")
        return OutOfOfficeSettings(
            state=cast(Literal["disabled", "enabled", "scheduled"], state),
            external_audience=cast(Literal["none", "known", "all"], audience),
            internal_reply=getattr(oof, "internal_reply", None) or None,
            external_reply=getattr(oof, "external_reply", None) or None,
            start=getattr(oof, "start", None),
            end=getattr(oof, "end", None),
        )

    def set_out_of_office(self, request: OutOfOfficeSettings) -> OutOfOfficeSettings:
        kwargs: dict[str, object] = {
            "state": _STATE_TO_EWS[request.state],
            "external_audience": _AUDIENCE_TO_EWS[request.external_audience],
            "internal_reply": request.internal_reply or "",
            "external_reply": request.external_reply or "",
        }
        if request.start is not None and request.end is not None:
            kwargs["start"] = self._to_ews_datetime(request.start)
            kwargs["end"] = self._to_ews_datetime(request.end)
        try:
            self.account.oof_settings = OofSettings(**kwargs)
        except Exception as exc:  # noqa: BLE001
            raise self._map_exception(exc) from exc
        return self.get_out_of_office()
