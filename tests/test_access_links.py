"""Tables a database only points at: MSysObjects rows of type 6.

What the engine writes for one was measured against DAO's
``TableDefs.Append``; see the linked-table gate.
"""

from __future__ import annotations

from pathlib import Path

import pytest

from pyopenvba.access import AccessDatabase
from pyopenvba.access.database import FLAG_LINKED, FLAG_LINKED_FOREIGN, OBJECT_LINKED_TABLE
from pyopenvba.access_read import AccessError

TEMPLATE = Path(__file__).parents[1] / "src" / "pyopenvba" / "_templates" / "blank_files" / "blank_database.accdb"
BACKEND = r"C:\share\backend.accdb"


def _fresh() -> AccessDatabase:
    return AccessDatabase(TEMPLATE)


def test_a_link_is_one_catalog_row_and_no_definition() -> None:
    db = _fresh()
    before = len(db.catalog())
    linked = db.link_table("Orders", BACKEND, "Orders")
    assert (linked.name, linked.database, linked.source, linked.connect) == ("Orders", BACKEND, "Orders", "")
    assert linked.is_jet
    assert len(db.catalog()) == before + 1
    entry = next(e for e in db.catalog() if e.name == "Orders")
    assert entry.type == OBJECT_LINKED_TABLE
    assert entry.flags == FLAG_LINKED
    # It is not a table of this file's own, so the table reader does not
    # offer it.
    assert "Orders" not in [t.name for t in db.tables()]
    with pytest.raises(AccessError, match="no table named"):
        db.table("Orders")


def test_a_foreign_source_keeps_its_prefix_and_flag() -> None:
    db = _fresh()
    linked = db.link_table("Rows", r"C:\share\csv", "rows#csv", connect="Text;")
    assert linked.connect == "Text;"
    assert not linked.is_jet
    entry = next(e for e in db.catalog() if e.name == "Rows")
    assert entry.flags == FLAG_LINKED | FLAG_LINKED_FOREIGN


def test_each_link_takes_the_next_id_up() -> None:
    db = _fresh()
    lowest = min(e.id for e in db.catalog())
    highest_negative = max(e.id for e in db.catalog() if e.id < 0)
    first = db.link_table("A", BACKEND, "One")
    second = db.link_table("B", BACKEND, "Two")
    assert (first.id, second.id) == (highest_negative + 1, highest_negative + 2)
    assert first.id > lowest


def test_a_link_gets_the_three_permission_rows() -> None:
    db = _fresh()
    linked = db.link_table("Orders", BACKEND, "Orders")
    aces = [r for r in db.table("MSysACEs").rows() if r["ObjectId"] == linked.id]
    assert len(aces) == 3
    assert {r["ACM"] for r in aces} == {0xFFEFF}


def test_a_name_already_in_use_is_refused() -> None:
    db = _fresh()
    db.link_table("Orders", BACKEND, "Orders")
    with pytest.raises(AccessError, match="already exists"):
        db.link_table("Orders", BACKEND, "Other")
    with pytest.raises(AccessError, match="already exists"):
        db.execute("CREATE TABLE Orders (A LONG)")


def test_dropping_a_link_leaves_nothing_behind() -> None:
    db = _fresh()
    linked = db.link_table("Orders", BACKEND, "Orders")
    db.drop_link("Orders")
    assert db.links() == []
    assert not [r for r in db.table("MSysACEs").rows() if r["ObjectId"] == linked.id]
    with pytest.raises(AccessError, match="no linked table named"):
        db.drop_link("Orders")


def test_links_survive_a_round_trip() -> None:
    db = _fresh()
    db.link_table("Orders", BACKEND, "Orders")
    db.link_table("Rows", r"C:\share\csv", "rows#csv", connect="Text;")
    again = AccessDatabase(db.to_bytes())
    assert [(x.name, x.database, x.source, x.connect) for x in again.links()] == [
        ("Orders", BACKEND, "Orders", ""),
        ("Rows", r"C:\share\csv", "rows#csv", "Text;"),
    ]
    assert again.link("orders").source == "Orders"
