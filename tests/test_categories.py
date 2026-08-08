from __future__ import annotations

from types import SimpleNamespace

import pytest
from exchangelib.errors import ErrorItemNotFound, UnauthorizedError

from outlook_mcp.errors import APIError
from outlook_mcp.exchange_client import EWSExchangeBackend
from outlook_mcp.models import ListCategoriesRequest

#: The document Outlook actually stores: categories namespace, colour as a
#: preset index, usageCount alongside.
_MASTER_XML = b"""<?xml version="1.0"?>
<categories default="Red" xmlns="CategoryList.xsd">
  <category name="Red category" color="0" usageCount="4" guid="{1}"/>
  <category name="Projects" color="7" usageCount="9" guid="{2}"/>
  <category name="No colour" color="-1" usageCount="9" guid="{3}"/>
  <category name="Bad colour" color="99" usageCount="not-a-number" guid="{4}"/>
  <category name="  " color="1" usageCount="1" guid="{5}"/>
</categories>
"""


def _backend_with_config(settings, config) -> EWSExchangeBackend:
    backend = EWSExchangeBackend(settings)

    def get_user_configuration(name):
        assert name == "CategoryList"
        if isinstance(config, Exception):
            raise config
        return config

    backend._account = SimpleNamespace(
        calendar=SimpleNamespace(get_user_configuration=get_user_configuration),
    )
    return backend


def test_list_categories_reads_the_mailbox_master_list(settings) -> None:
    backend = _backend_with_config(settings, SimpleNamespace(xml_data=_MASTER_XML))

    result = backend.list_categories(ListCategoriesRequest())

    by_name = {usage.name: usage for usage in result}
    assert by_name["Red category"].color == "red"
    assert by_name["Red category"].count == 4
    assert by_name["Projects"].color == "blue"
    # -1 and out-of-range indexes mean "no colour", not a crash.
    assert by_name["No colour"].color is None
    assert by_name["Bad colour"].color is None
    assert by_name["Bad colour"].count == 0
    # Blank names are Outlook artifacts, not categories.
    assert "  " not in by_name


def test_list_categories_sorts_by_count_then_name(settings) -> None:
    backend = _backend_with_config(settings, SimpleNamespace(xml_data=_MASTER_XML))

    result = backend.list_categories(ListCategoriesRequest())

    assert [usage.name for usage in result] == [
        "No colour",
        "Projects",
        "Red category",
        "Bad colour",
    ]


def test_list_categories_falls_back_to_scanning_when_no_master_list(settings) -> None:
    """A mailbox that has never defined categories has no CategoryList item.
    The recent-messages heuristic still answers, with colours unknown."""

    class FakeQuerySet:
        def __init__(self, items):
            self.items = items

        def all(self):
            return self

        def only(self, *fields):
            return self

        def order_by(self, *fields):
            return self

        def __getitem__(self, item):
            return self.items

    message = SimpleNamespace(categories=["Ops", "ops", "Later"])
    backend = _backend_with_config(settings, ErrorItemNotFound("no such item"))
    folder = FakeQuerySet([message, SimpleNamespace(categories=["Ops"])])
    backend._resolve_folder = lambda value: folder  # type: ignore[method-assign]

    result = backend.list_categories(ListCategoriesRequest(folders=["inbox"]))

    assert [(usage.name, usage.count) for usage in result] == [("Ops", 3), ("Later", 1)]
    assert all(usage.color is None for usage in result)


def test_list_categories_falls_back_on_unparseable_xml(settings) -> None:
    backend = _backend_with_config(settings, SimpleNamespace(xml_data=b"<not xml"))
    backend._resolve_folder = lambda value: (_ for _ in ()).throw(AssertionError)  # type: ignore

    # No folders scanned in this test: an empty fallback pass is fine, the point
    # is that broken XML does not raise.
    request = ListCategoriesRequest(folders=["inbox"])

    class EmptyFolder:
        def all(self):
            return self

        def only(self, *f):
            return self

        def order_by(self, *f):
            return self

        def __getitem__(self, item):
            return []

    backend._resolve_folder = lambda value: EmptyFolder()  # type: ignore[method-assign]
    assert backend.list_categories(request) == []


def test_list_categories_propagates_auth_failures_instead_of_falling_back(settings) -> None:
    """Only "the config item does not exist" means fallback. An auth failure says
    nothing about the master list and must surface as itself."""
    backend = _backend_with_config(settings, UnauthorizedError("bad credentials"))

    with pytest.raises(APIError) as excinfo:
        backend.list_categories(ListCategoriesRequest())

    assert excinfo.value.code == "auth_failed"
