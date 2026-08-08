from __future__ import annotations

from types import SimpleNamespace

import pytest
from exchangelib.properties import Actions, Address, Conditions, Rule

import outlook_mcp.exchange_client.mailbox as mailbox_module
from outlook_mcp.errors import NotFoundError
from outlook_mcp.exchange_client import EWSExchangeBackend
from outlook_mcp.models import (
    CreateInboxRuleActions,
    CreateInboxRuleConditions,
    InboxRule,
    CreateInboxRuleRequest,
    DeleteInboxRuleRequest,
    InboxRuleActions,
    InboxRuleConditions,
    UpdateInboxRuleRequest,
)


class FakeRuleService:
    """Stands in for Get/Create/Set/DeleteInboxRule; records calls per class."""

    calls: dict[str, list] = {}
    rules: list = []

    def __init__(self, account) -> None:
        pass

    @classmethod
    def reset(cls, rules=None) -> None:
        cls.calls = {"get": [], "create": [], "set": [], "delete": []}
        cls.rules = rules or []


class FakeGet(FakeRuleService):
    def call(self):
        FakeRuleService.calls["get"].append(True)
        yield from FakeRuleService.rules


class FakeCreate(FakeRuleService):
    def call(self, rule, remove_outlook_rule_blob=True):
        def lazy():
            # Like the real service: the request is only sent when consumed.
            FakeRuleService.calls["create"].append((rule, remove_outlook_rule_blob))
            yield from ()

        return lazy()


class FakeSet(FakeRuleService):
    def call(self, rule, remove_outlook_rule_blob=True):
        def lazy():
            FakeRuleService.calls["set"].append((rule, remove_outlook_rule_blob))
            yield from ()

        return lazy()


class FakeDelete(FakeRuleService):
    def call(self, rule, remove_outlook_rule_blob=True):
        def lazy():
            FakeRuleService.calls["delete"].append((rule.id, remove_outlook_rule_blob))
            yield from ()

        return lazy()


@pytest.fixture
def backend(settings, monkeypatch) -> EWSExchangeBackend:
    monkeypatch.setattr(mailbox_module, "GetInboxRules", FakeGet)
    monkeypatch.setattr(mailbox_module, "CreateInboxRule", FakeCreate)
    monkeypatch.setattr(mailbox_module, "SetInboxRule", FakeSet)
    monkeypatch.setattr(mailbox_module, "DeleteInboxRule", FakeDelete)
    FakeRuleService.reset()
    result = EWSExchangeBackend(settings)
    result._account = SimpleNamespace()
    return result


def _server_rule(**overrides) -> Rule:
    fields = dict(
        id="rule-1",
        display_name="From boss",
        priority=1,
        is_enabled=True,
        conditions=Conditions(
            from_addresses=[Address(email_address="boss@example.com")],
            importance="High",
        ),
        actions=Actions(mark_as_read=True),
    )
    fields.update(overrides)
    return Rule(**fields)


def test_list_inbox_rules_maps_the_curated_subset(backend) -> None:
    FakeRuleService.reset(rules=[_server_rule()])

    rules = backend.list_inbox_rules()

    assert len(rules) == 1
    rule = rules[0]
    assert rule.display_name == "From boss"
    assert rule.conditions.from_addresses == ["boss@example.com"]
    assert rule.conditions.importance == "high"
    assert rule.actions.mark_as_read is True


def test_create_inbox_rule_builds_ews_conditions_and_actions(backend) -> None:
    request = CreateInboxRuleRequest(
        display_name="Filed",
        conditions=CreateInboxRuleConditions(from_addresses=["boss@example.com"]),
        actions=CreateInboxRuleActions(assign_categories=["Boss"], mark_as_read=True),
    )
    FakeRuleService.reset(
        rules=[_server_rule(display_name="Filed", actions=Actions(mark_as_read=True))]
    )

    created = backend.create_inbox_rule(request)

    sent_rule, blob = FakeRuleService.calls["create"][0]
    assert sent_rule.display_name == "Filed"
    assert sent_rule.conditions.from_addresses[0].email_address == "boss@example.com"
    assert sent_rule.actions.assign_categories == ["Boss"]
    assert blob is True
    # The created rule is re-read from the server so the caller gets its id.
    assert created.id == "rule-1"


def test_update_inbox_rule_sends_the_servers_own_rule_object_back(backend) -> None:
    """Round-tripping through the curated model would drop any condition the
    model does not cover; the server's Rule object must go back as-is with only
    the requested fields changed."""
    server_rule = _server_rule()
    FakeRuleService.reset(rules=[server_rule])

    result = backend.update_inbox_rule(UpdateInboxRuleRequest(id="rule-1", is_enabled=False))

    sent_rule, _ = FakeRuleService.calls["set"][0]
    assert sent_rule is server_rule
    assert sent_rule.is_enabled is False
    # Untouched parts survive.
    assert sent_rule.conditions.from_addresses[0].email_address == "boss@example.com"
    assert result.is_enabled is False


def test_update_inbox_rule_unknown_id_is_not_found(backend) -> None:
    FakeRuleService.reset(rules=[_server_rule()])

    with pytest.raises(NotFoundError):
        backend.update_inbox_rule(UpdateInboxRuleRequest(id="nope", is_enabled=False))


def test_delete_inbox_rule_deletes_by_id(backend) -> None:
    result = backend.delete_inbox_rule(DeleteInboxRuleRequest(id="rule-1"))

    assert FakeRuleService.calls["delete"] == [("rule-1", True)]
    assert result.status == "deleted"


def test_create_inbox_rule_requires_a_condition_and_an_action() -> None:
    with pytest.raises(ValueError, match="at least one condition"):
        CreateInboxRuleRequest(
            display_name="Empty",
            conditions=CreateInboxRuleConditions(),
            actions=CreateInboxRuleActions(mark_as_read=True),
        )

    with pytest.raises(ValueError, match="at least one action"):
        CreateInboxRuleRequest(
            display_name="Empty",
            conditions=CreateInboxRuleConditions(has_attachments=True),
            actions=CreateInboxRuleActions(),
        )


def test_update_inbox_rule_requires_something_to_update() -> None:
    with pytest.raises(ValueError, match="nothing to update"):
        UpdateInboxRuleRequest(id="rule-1")


def test_a_rule_read_from_the_server_survives_x500_addresses() -> None:
    """A live mailbox had a rule whose from_addresses held an X.500 distinguished
    name, which crashed list_inbox_rules. Read models must accept whatever the
    server reports; only *creating* a rule stays strict."""
    x500 = "/o=TANDER/ou=Exchange Administrative Group/cn=Recipients/cn=abc-support"

    rule = InboxRule(
        display_name="From support notebook",
        conditions=InboxRuleConditions(from_addresses=[x500]),
        actions=InboxRuleActions(forward_to=[x500]),
    )

    assert rule.conditions.from_addresses == [x500]

    with pytest.raises(ValueError):
        CreateInboxRuleConditions(from_addresses=[x500])
    with pytest.raises(ValueError):
        CreateInboxRuleActions(forward_to=[x500])
