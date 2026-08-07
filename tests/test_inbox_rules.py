from __future__ import annotations

from types import SimpleNamespace

import pytest
from exchangelib.properties import Actions, Address, Conditions, FolderId, MoveToFolder, Rule

from outlook_mcp.exchange_client import EWSExchangeBackend
from outlook_mcp.models import (
    ActionResult,
    CreateRuleRequest,
    DeleteRuleRequest,
    MailRule,
    UpdateRuleRequest,
)


def _backend() -> EWSExchangeBackend:
    backend = EWSExchangeBackend.__new__(EWSExchangeBackend)
    backend.settings = SimpleNamespace(exchange_timezone_fallback=None)
    return backend


def test_list_rules_maps_conditions_and_actions() -> None:
    backend = _backend()
    ews_rule = Rule(
        id="rule-1",
        display_name="Move newsletters",
        priority=2,
        is_enabled=True,
        conditions=Conditions(
            from_addresses=[Address(email_address="news@example.com")],
            contains_subject_strings=["sale"],
            has_attachments=False,
        ),
        actions=Actions(
            move_to_folder=MoveToFolder(folder_id=FolderId(id="folder-newsletters")),
            mark_as_read=True,
            assign_categories=["Newsletter"],
            stop_processing_rules=True,
        ),
    )
    backend._account = SimpleNamespace(rules=[ews_rule])

    result = backend.list_rules()

    assert result == [
        MailRule(
            id="rule-1",
            display_name="Move newsletters",
            priority=2,
            is_enabled=True,
            from_addresses=["news@example.com"],
            contains_subject_strings=["sale"],
            has_attachments=False,
            move_to_folder="folder-newsletters",
            mark_as_read=True,
            assign_categories=["Newsletter"],
            delete=False,
            stop_processing_rules=True,
        )
    ]


def test_list_rules_handles_rule_with_no_conditions() -> None:
    backend = _backend()
    ews_rule = Rule(
        id="rule-2",
        display_name="Delete spam",
        priority=1,
        is_enabled=True,
        conditions=None,
        actions=Actions(delete=True, stop_processing_rules=True),
    )
    backend._account = SimpleNamespace(rules=[ews_rule])

    result = backend.list_rules()

    assert result[0].from_addresses == []
    assert result[0].delete is True


def test_create_rule_builds_conditions_and_actions_and_saves() -> None:
    backend = _backend()
    events: list[tuple] = []

    class FakeFolder:
        id = "folder-archive"

    def create_rule(rule):
        events.append(("create_rule", rule.display_name, rule.priority))
        rule.id = "rule-new"

    backend._account = SimpleNamespace(create_rule=create_rule)
    backend._resolve_folder = lambda value: FakeFolder()

    result = backend.create_rule(
        CreateRuleRequest(
            display_name="Archive receipts",
            priority=3,
            from_addresses=["receipts@example.com"],
            contains_subject_strings=["receipt"],
            move_to_folder="archive",
        )
    )

    assert events == [("create_rule", "Archive receipts", 3)]
    assert result == ActionResult(id="rule-new", status="created")


def test_create_rule_requires_at_least_one_action() -> None:
    with pytest.raises(Exception):
        CreateRuleRequest(display_name="No-op rule")


def test_update_rule_replaces_the_whole_rule() -> None:
    backend = _backend()
    events: list[tuple] = []

    def set_rule(rule):
        events.append(("set_rule", rule.id, rule.display_name))

    backend._account = SimpleNamespace(set_rule=set_rule)

    result = backend.update_rule(
        UpdateRuleRequest(id="rule-1", display_name="Renamed rule", mark_as_read=True)
    )

    assert events == [("set_rule", "rule-1", "Renamed rule")]
    assert result == ActionResult(id="rule-1", status="updated")


def test_delete_rule_only_needs_the_id() -> None:
    backend = _backend()
    events: list[tuple] = []

    def delete_rule(rule):
        events.append(("delete_rule", rule.id))

    backend._account = SimpleNamespace(delete_rule=delete_rule)

    result = backend.delete_rule(DeleteRuleRequest(id="rule-1"))

    assert events == [("delete_rule", "rule-1")]
    assert result == ActionResult(id="rule-1", status="deleted")
