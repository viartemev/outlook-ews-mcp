from __future__ import annotations

from types import SimpleNamespace

import pytest
from exchangelib.errors import ErrorAccessDenied, UnauthorizedError

import outlook_mcp.exchange_client.mailbox as mailbox_module
from outlook_mcp.errors import APIError, NotFoundError
from outlook_mcp.exchange_client import EWSExchangeBackend
from outlook_mcp.models import (
    CreateInboxRuleActions,
    CreateInboxRuleConditions,
    CreateInboxRuleRequest,
    DeleteInboxRuleRequest,
    OutOfOfficeSettings,
    UpdateInboxRuleRequest,
)


class _RaisingAccount:
    """An account whose delegate/OOF surfaces fail like a locked-down mailbox."""

    def __init__(self, error: Exception) -> None:
        self._error = error

    @property
    def delegates(self):
        raise self._error

    @property
    def oof_settings(self):
        raise self._error

    @oof_settings.setter
    def oof_settings(self, value):
        raise self._error


def _backend(settings, account) -> EWSExchangeBackend:
    backend = EWSExchangeBackend(settings)
    backend._account = account
    return backend


def test_list_delegates_maps_access_denied(settings) -> None:
    backend = _backend(settings, _RaisingAccount(ErrorAccessDenied("no rights")))

    with pytest.raises(APIError) as excinfo:
        backend.list_delegates()

    assert excinfo.value.code == "permission_denied"


def test_get_out_of_office_maps_auth_failure(settings) -> None:
    backend = _backend(settings, _RaisingAccount(UnauthorizedError("bad creds")))

    with pytest.raises(APIError) as excinfo:
        backend.get_out_of_office()

    assert excinfo.value.code == "auth_failed"


def test_set_out_of_office_maps_write_failure(settings) -> None:
    backend = _backend(settings, _RaisingAccount(ErrorAccessDenied("read only")))

    with pytest.raises(APIError) as excinfo:
        backend.set_out_of_office(OutOfOfficeSettings(state="disabled"))

    assert excinfo.value.code == "permission_denied"


class _Service:
    """One fake for all four rule services; behaviour injected per test."""

    get_result: object = ()
    write_result: object = ()

    def __init__(self, account) -> None:
        pass

    def call(self, *args, **kwargs):
        result = self.get_result if not (args or kwargs) else self.write_result
        if isinstance(result, Exception):
            raise result
        yield from result


def _rules_backend(settings, monkeypatch, get_result=(), write_result=()) -> EWSExchangeBackend:
    _Service.get_result = get_result
    _Service.write_result = write_result
    for name in ("GetInboxRules", "CreateInboxRule", "SetInboxRule", "DeleteInboxRule"):
        monkeypatch.setattr(mailbox_module, name, _Service)
    backend = EWSExchangeBackend(settings)
    backend._account = SimpleNamespace()
    return backend


def _server_rule(rule_id="rule-1", name="Existing"):
    return SimpleNamespace(
        id=rule_id,
        display_name=name,
        priority=1,
        is_enabled=True,
        is_not_supported=False,
        conditions=None,
        actions=None,
    )


def _create_request(name="My rule") -> CreateInboxRuleRequest:
    return CreateInboxRuleRequest(
        display_name=name,
        conditions=CreateInboxRuleConditions(has_attachments=True),
        actions=CreateInboxRuleActions(mark_as_read=True),
    )


def test_list_inbox_rules_maps_a_service_failure(settings, monkeypatch) -> None:
    backend = _rules_backend(settings, monkeypatch, get_result=UnauthorizedError("bad creds"))

    with pytest.raises(APIError) as excinfo:
        backend.list_inbox_rules()

    assert excinfo.value.code == "auth_failed"


def test_list_inbox_rules_raises_on_an_error_element(settings, monkeypatch) -> None:
    """GetInboxRules yields per-rule elements; any one may be an Exception."""
    backend = _rules_backend(settings, monkeypatch, get_result=[ErrorAccessDenied("rule blocked")])

    with pytest.raises(APIError) as excinfo:
        backend.list_inbox_rules()

    assert excinfo.value.code == "permission_denied"


def test_list_inbox_rules_skips_none_elements(settings, monkeypatch) -> None:
    backend = _rules_backend(settings, monkeypatch, get_result=[None, _server_rule()])

    result = backend.list_inbox_rules()

    assert [rule.display_name for rule in result] == ["Existing"]


def test_create_inbox_rule_maps_a_service_failure(settings, monkeypatch) -> None:
    backend = _rules_backend(settings, monkeypatch, write_result=ErrorAccessDenied("nope"))

    with pytest.raises(APIError) as excinfo:
        backend.create_inbox_rule(_create_request())

    assert excinfo.value.code == "permission_denied"


def test_create_inbox_rule_surfaces_a_rule_operation_error_element(settings, monkeypatch) -> None:
    backend = _rules_backend(
        settings, monkeypatch, write_result=[ErrorAccessDenied("operation error")]
    )

    with pytest.raises(APIError) as excinfo:
        backend.create_inbox_rule(_create_request())

    assert excinfo.value.code == "permission_denied"


def test_create_inbox_rule_falls_back_to_the_local_model_when_relist_misses_it(
    settings, monkeypatch
) -> None:
    """Some servers delay rule visibility; the answer is still the created rule."""
    backend = _rules_backend(settings, monkeypatch, get_result=[_server_rule()])

    result = backend.create_inbox_rule(_create_request(name="Not listed yet"))

    assert result.display_name == "Not listed yet"
    assert result.id is None


def test_update_inbox_rule_maps_a_lookup_failure(settings, monkeypatch) -> None:
    backend = _rules_backend(settings, monkeypatch, get_result=UnauthorizedError("bad creds"))

    with pytest.raises(APIError) as excinfo:
        backend.update_inbox_rule(UpdateInboxRuleRequest(id="rule-1", is_enabled=False))

    assert excinfo.value.code == "auth_failed"


def test_update_inbox_rule_maps_a_write_failure(settings, monkeypatch) -> None:
    backend = _rules_backend(
        settings,
        monkeypatch,
        get_result=[_server_rule()],
        write_result=ErrorAccessDenied("nope"),
    )

    with pytest.raises(APIError) as excinfo:
        backend.update_inbox_rule(UpdateInboxRuleRequest(id="rule-1", priority=2))

    assert excinfo.value.code == "permission_denied"


def test_update_inbox_rule_reports_an_unknown_id_as_not_found(settings, monkeypatch) -> None:
    backend = _rules_backend(settings, monkeypatch, get_result=[_server_rule()])

    with pytest.raises(NotFoundError):
        backend.update_inbox_rule(UpdateInboxRuleRequest(id="no-such-rule", is_enabled=False))


def test_delete_inbox_rule_maps_a_service_failure(settings, monkeypatch) -> None:
    backend = _rules_backend(settings, monkeypatch, write_result=ErrorAccessDenied("nope"))

    with pytest.raises(APIError) as excinfo:
        backend.delete_inbox_rule(DeleteInboxRuleRequest(id="rule-1"))

    assert excinfo.value.code == "permission_denied"
