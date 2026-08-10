from __future__ import annotations

from types import SimpleNamespace

from outlook_mcp.exchange_client import EWSExchangeBackend
from outlook_mcp.models import ListRoomsRequest, RoomInfo, RoomListInfo


def _backend() -> EWSExchangeBackend:
    backend = EWSExchangeBackend.__new__(EWSExchangeBackend)
    backend.settings = SimpleNamespace(exchange_timezone_fallback=None)
    return backend


def test_list_room_lists_maps_name_and_email() -> None:
    backend = _backend()
    fake_room_list = SimpleNamespace(name="Building A", email_address="buildinga@example.com")
    backend._account = SimpleNamespace(
        protocol=SimpleNamespace(get_roomlists=lambda: [fake_room_list])
    )

    result = backend.list_room_lists()

    assert result == [RoomListInfo(name="Building A", email="buildinga@example.com")]


def test_list_rooms_calls_get_rooms_with_the_requested_list() -> None:
    backend = _backend()
    calls: list[str] = []
    fake_room = SimpleNamespace(name="Room 101", email_address="room101@example.com")

    def get_rooms(room_list):
        calls.append(room_list)
        return [fake_room]

    backend._account = SimpleNamespace(protocol=SimpleNamespace(get_rooms=get_rooms))

    result = backend.list_rooms(ListRoomsRequest(room_list="buildinga@example.com"))

    assert calls == ["buildinga@example.com"]
    assert result == [RoomInfo(name="Room 101", email="room101@example.com")]
