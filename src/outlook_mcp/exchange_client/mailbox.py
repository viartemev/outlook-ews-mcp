from __future__ import annotations

from typing import Literal, cast

from exchangelib.properties import (
    Actions,
    Address,
    Conditions,
    FolderId,
    MoveToFolder,
    Rule,
)
from exchangelib.services import (
    CreateInboxRule,
    DeleteInboxRule,
    GetInboxRules,
    SetInboxRule,
)
from exchangelib.settings import OofSettings

from ..errors import NotFoundError
from ..models import (
    ActionResult,
    CreateInboxRuleRequest,
    DelegateInfo,
    DelegatePermissionLevels,
    DeleteInboxRuleRequest,
    InboxRule,
    InboxRuleActions,
    InboxRuleConditions,
    OutOfOfficeSettings,
    UpdateInboxRuleRequest,
)
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

    def _rule_to_model(self, rule: Rule) -> InboxRule:
        conditions = getattr(rule, "conditions", None)
        actions = getattr(rule, "actions", None)
        move_to = getattr(actions, "move_to_folder", None)
        folder_id = getattr(getattr(move_to, "folder_id", None), "id", None)
        importance = getattr(conditions, "importance", None)
        return InboxRule(
            id=getattr(rule, "id", None),
            display_name=getattr(rule, "display_name", "") or "",
            priority=getattr(rule, "priority", 1) or 1,
            is_enabled=bool(getattr(rule, "is_enabled", False)),
            is_not_supported=bool(getattr(rule, "is_not_supported", False)),
            conditions=InboxRuleConditions(
                from_addresses=[
                    address.email_address
                    for address in getattr(conditions, "from_addresses", None) or []
                    if getattr(address, "email_address", None)
                ],
                contains_sender_strings=list(
                    getattr(conditions, "contains_sender_strings", None) or []
                ),
                contains_subject_strings=list(
                    getattr(conditions, "contains_subject_strings", None) or []
                ),
                contains_subject_or_body_strings=list(
                    getattr(conditions, "contains_subject_or_body_strings", None) or []
                ),
                has_attachments=getattr(conditions, "has_attachments", None),
                importance=importance.lower() if importance else None,
            ),
            actions=InboxRuleActions(
                move_to_folder=folder_id,
                assign_categories=list(getattr(actions, "assign_categories", None) or []),
                mark_as_read=getattr(actions, "mark_as_read", None),
                delete=getattr(actions, "delete", None),
                forward_to=[
                    address.email_address
                    for address in getattr(actions, "forward_to_recipients", None) or []
                    if getattr(address, "email_address", None)
                ],
            ),
        )

    def _rule_from_request(self, request: CreateInboxRuleRequest) -> Rule:
        conditions = Conditions(
            from_addresses=[
                Address(email_address=str(address)) for address in request.conditions.from_addresses
            ]
            or None,
            contains_sender_strings=request.conditions.contains_sender_strings or None,
            contains_subject_strings=request.conditions.contains_subject_strings or None,
            contains_subject_or_body_strings=(
                request.conditions.contains_subject_or_body_strings or None
            ),
            has_attachments=request.conditions.has_attachments,
            importance=(
                request.conditions.importance.capitalize()
                if request.conditions.importance
                else None
            ),
        )
        actions_kwargs: dict[str, object] = {
            "assign_categories": request.actions.assign_categories or None,
            "mark_as_read": request.actions.mark_as_read,
            "delete": request.actions.delete,
            "forward_to_recipients": [
                Address(email_address=str(address)) for address in request.actions.forward_to
            ]
            or None,
        }
        if request.actions.move_to_folder:
            folder = self._resolve_folder(request.actions.move_to_folder)
            actions_kwargs["move_to_folder"] = MoveToFolder(
                folder_id=FolderId(id=folder.id, changekey=getattr(folder, "changekey", None))
            )
        return Rule(
            display_name=request.display_name,
            priority=request.priority,
            is_enabled=request.is_enabled,
            conditions=conditions,
            actions=Actions(**actions_kwargs),
        )

    def list_inbox_rules(self) -> list[InboxRule]:
        try:
            rules = list(GetInboxRules(account=self.account).call())
        except Exception as exc:  # noqa: BLE001
            raise self._map_exception(exc) from exc
        result: list[InboxRule] = []
        for rule in rules:
            if isinstance(rule, Exception):
                raise self._map_exception(rule)
            if rule is None:
                continue
            result.append(self._rule_to_model(rule))
        return result

    def create_inbox_rule(self, request: CreateInboxRuleRequest) -> InboxRule:
        rule = self._rule_from_request(request)
        try:
            CreateInboxRule(account=self.account).call(
                rule=rule, remove_outlook_rule_blob=request.remove_outlook_rule_blob
            )
        except Exception as exc:  # noqa: BLE001
            raise self._map_exception(exc) from exc
        # CreateInboxRule returns no id; re-list and match on the unique name.
        for created in self.list_inbox_rules():
            if created.display_name == request.display_name:
                return created
        return self._rule_to_model(rule)

    def update_inbox_rule(self, request: UpdateInboxRuleRequest) -> InboxRule:
        try:
            rules = list(GetInboxRules(account=self.account).call())
        except Exception as exc:  # noqa: BLE001
            raise self._map_exception(exc) from exc
        for rule in rules:
            if isinstance(rule, Exception) or rule is None:
                continue
            if getattr(rule, "id", None) == request.id:
                # Send the server's own Rule object back with only the two fields
                # changed: round-tripping through this API's curated subset would
                # silently drop conditions/actions the subset does not model.
                if request.is_enabled is not None:
                    rule.is_enabled = request.is_enabled
                if request.priority is not None:
                    rule.priority = request.priority
                try:
                    SetInboxRule(account=self.account).call(
                        rule=rule,
                        remove_outlook_rule_blob=request.remove_outlook_rule_blob,
                    )
                except Exception as exc:  # noqa: BLE001
                    raise self._map_exception(exc) from exc
                return self._rule_to_model(rule)
        raise NotFoundError(request.id)

    def delete_inbox_rule(self, request: DeleteInboxRuleRequest) -> ActionResult:
        rule = Rule(id=request.id)
        try:
            DeleteInboxRule(account=self.account).call(
                rule=rule, remove_outlook_rule_blob=request.remove_outlook_rule_blob
            )
        except Exception as exc:  # noqa: BLE001
            raise self._map_exception(exc) from exc
        return ActionResult(id=request.id, status="deleted")
