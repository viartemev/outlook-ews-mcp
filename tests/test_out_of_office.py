from __future__ import annotations

from types import SimpleNamespace

import pytest
from exchangelib.ewsdatetime import EWSDateTime, EWSTimeZone
from exchangelib.settings import OofSettings

from outlook_mcp.exchange_client import EWSExchangeBackend
from outlook_mcp.models import OutOfOfficeSettings


def _backend_with_oof(settings, oof) -> EWSExchangeBackend:
    backend = EWSExchangeBackend(settings)
    backend._account = SimpleNamespace(
        oof_settings=oof, default_timezone=EWSTimeZone("Europe/Moscow")
    )
    return backend


def test_get_out_of_office_maps_ews_values_to_lowercase(settings) -> None:
    backend = _backend_with_oof(
        settings,
        OofSettings(
            state="Scheduled",
            external_audience="Known",
            internal_reply="I am away",
            external_reply="Out of office",
            start=EWSDateTime(2026, 8, 10, 9, 0, tzinfo=EWSTimeZone("UTC")),
            end=EWSDateTime(2026, 8, 20, 18, 0, tzinfo=EWSTimeZone("UTC")),
        ),
    )

    result = backend.get_out_of_office()

    assert result.state == "scheduled"
    assert result.external_audience == "known"
    assert result.internal_reply == "I am away"
    assert result.start is not None and result.end is not None


def test_set_out_of_office_maps_lowercase_to_ews_and_reads_back(settings) -> None:
    backend = _backend_with_oof(settings, OofSettings(state="Disabled"))

    result = backend.set_out_of_office(
        OutOfOfficeSettings(
            state="enabled",
            external_audience="none",
            internal_reply="Back Monday",
            external_reply="Back Monday",
        )
    )

    stored = backend._account.oof_settings
    assert stored.state == "Enabled"
    assert stored.external_audience == "None"
    assert result.state == "enabled"
    assert result.internal_reply == "Back Monday"


def test_set_out_of_office_scheduled_sends_mailbox_local_datetimes(settings) -> None:
    backend = _backend_with_oof(settings, OofSettings(state="Disabled"))

    backend.set_out_of_office(
        OutOfOfficeSettings.model_validate(
            {
                "state": "scheduled",
                "internal_reply": "Away",
                "start": "2026-08-10T09:00:00+00:00",
                "end": "2026-08-20T18:00:00+00:00",
            }
        )
    )

    stored = backend._account.oof_settings
    assert stored.state == "Scheduled"
    assert isinstance(stored.start, EWSDateTime)
    assert stored.start.tzinfo.key == "Europe/Moscow"


def test_out_of_office_scheduled_requires_both_dates() -> None:
    with pytest.raises(ValueError, match="requires both start and end"):
        OutOfOfficeSettings.model_validate({"state": "scheduled", "internal_reply": "Away"})


def test_out_of_office_scheduled_rejects_a_reversed_window() -> None:
    with pytest.raises(ValueError, match="end must be greater than start"):
        OutOfOfficeSettings.model_validate(
            {
                "state": "scheduled",
                "start": "2026-08-20T18:00:00+00:00",
                "end": "2026-08-10T09:00:00+00:00",
            }
        )


def test_list_delegates_maps_user_and_permission_fields(settings) -> None:
    from outlook_mcp.exchange_client import EWSExchangeBackend as Backend

    delegate = SimpleNamespace(
        user_id=SimpleNamespace(
            primary_smtp_address="assistant@example.com", display_name="Assistant"
        ),
        delegate_permissions=SimpleNamespace(
            calendar_folder_permission_level="Editor",
            inbox_folder_permission_level="Reviewer",
            tasks_folder_permission_level=None,
            contacts_folder_permission_level="None",
        ),
        receive_copies_of_meeting_messages=True,
        view_private_items=False,
    )
    backend = Backend(settings)
    backend._account = SimpleNamespace(delegates=[delegate])

    result = backend.list_delegates()

    assert len(result) == 1
    info = result[0]
    assert info.email == "assistant@example.com"
    assert info.permissions.calendar == "Editor"
    assert info.permissions.inbox == "Reviewer"
    # A level the server reports as absent reads as "None", not a crash.
    assert info.permissions.tasks == "None"
    assert info.receives_copies_of_meeting_messages is True


def test_list_delegates_empty_mailbox_is_a_valid_answer(settings) -> None:
    from outlook_mcp.exchange_client import EWSExchangeBackend as Backend

    backend = Backend(settings)
    backend._account = SimpleNamespace(delegates=[])

    assert backend.list_delegates() == []
