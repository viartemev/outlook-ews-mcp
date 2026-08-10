from __future__ import annotations

from datetime import UTC, datetime
from types import SimpleNamespace

import pytest
from exchangelib.settings import OofSettings

from outlook_mcp.exchange_client import EWSExchangeBackend
from outlook_mcp.models import ActionResult, OofSettingsModel, SetOofSettingsRequest


def _backend() -> EWSExchangeBackend:
    backend = EWSExchangeBackend.__new__(EWSExchangeBackend)
    backend.settings = SimpleNamespace(exchange_timezone_fallback=None)
    return backend


def test_get_oof_settings_maps_disabled_state() -> None:
    backend = _backend()
    backend._account = SimpleNamespace(
        oof_settings=OofSettings(state="Disabled", external_audience="All")
    )

    result = backend.get_oof_settings()

    assert result == OofSettingsModel(state="disabled", external_audience="all")


def test_get_oof_settings_maps_scheduled_state_with_replies() -> None:
    backend = _backend()
    start = datetime(2026, 6, 1, 9, 0, tzinfo=UTC)
    end = datetime(2026, 6, 5, 18, 0, tzinfo=UTC)
    backend._account = SimpleNamespace(
        oof_settings=OofSettings(
            state="Scheduled",
            external_audience="Known",
            start=start,
            end=end,
            internal_reply="Back Monday",
            external_reply="Away",
        )
    )

    result = backend.get_oof_settings()

    assert result.state == "scheduled"
    assert result.external_audience == "known"
    assert result.internal_reply == "Back Monday"
    assert result.external_reply == "Away"


def test_set_oof_settings_builds_ews_settings_and_assigns() -> None:
    backend = _backend()
    captured: list[OofSettings] = []

    class FakeAccount:
        @property
        def oof_settings(self):
            raise AssertionError("not read in this test")

        @oof_settings.setter
        def oof_settings(self, value):
            captured.append(value)

    backend._account = FakeAccount()

    result = backend.set_oof_settings(
        SetOofSettingsRequest(
            state="enabled",
            external_audience="known",
            internal_reply="I'm out",
            external_reply="I'm out",
        )
    )

    assert len(captured) == 1
    assert captured[0].state == "Enabled"
    assert captured[0].external_audience == "Known"
    assert result == ActionResult(id="oof", status="updated")


def test_set_oof_settings_requires_replies_unless_disabled() -> None:
    with pytest.raises(Exception):
        SetOofSettingsRequest(state="enabled")


def test_set_oof_settings_requires_window_when_scheduled() -> None:
    with pytest.raises(Exception):
        SetOofSettingsRequest(state="scheduled", internal_reply="x", external_reply="x")


def test_set_oof_settings_disabled_needs_no_replies() -> None:
    request = SetOofSettingsRequest(state="disabled")
    assert request.internal_reply is None
