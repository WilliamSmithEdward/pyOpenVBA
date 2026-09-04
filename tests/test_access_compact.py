"""Giving back the free pages a file ends with.

This is not Access's Compact and Repair, which rebuilds the database and
moves every page.  It reclaims the run of free pages at the end, which is
what a dropped table or a large delete leaves behind, and moves nothing.
"""

from __future__ import annotations

from pathlib import Path

import pytest

from pyopenvba.access import AccessDatabase, ColumnSpec, IndexSpec
from pyopenvba.access._pages import PAGE_SIZE


def filled(path: Path, rows: int = 2000) -> AccessDatabase:
    """A database with a small table to keep and a large one to drop."""
    db = AccessDatabase.create_new(path)
    keep = db.create_table(
        "Keep",
        [ColumnSpec("Id", "Long"), ColumnSpec("Note", "Text", size=60)],
        [IndexSpec("PrimaryKey", ("Id",), primary=True)],
    )
    for number in range(20):
        keep.insert_row({"Id": number, "Note": f"row {number}"})
    temp = db.create_table(
        "Temp",
        [ColumnSpec("Id", "Long"), ColumnSpec("Pad", "Text", size=200)],
        [IndexSpec("TempKey", ("Id",), primary=True)],
    )
    for number in range(rows):
        temp.insert_row({"Id": number, "Pad": "x" * 180})
    return db


@pytest.fixture
def db(tmp_path: Path) -> AccessDatabase:
    path = tmp_path / "orders.accdb"
    database = filled(path)
    database.save(path)
    return AccessDatabase(path)


def test_a_dropped_table_leaves_pages_compact_gives_back(db: AccessDatabase) -> None:
    before = db.store.page_count
    db.drop_table("Temp")

    reclaimed = db.compact()

    assert reclaimed > 0
    assert db.store.page_count == before - reclaimed
    assert len(db.to_bytes()) == db.store.page_count * PAGE_SIZE


def test_what_is_kept_survives(db: AccessDatabase, tmp_path: Path) -> None:
    db.drop_table("Temp")
    db.compact()
    out = tmp_path / "compacted.accdb"
    db.save(out)

    reopened = AccessDatabase(out)
    assert reopened.table_names() == ["Keep"]
    rows = list(reopened.table("Keep").rows())
    assert len(rows) == 20
    assert rows[0]["Note"] == "row 0"


def test_the_file_still_takes_writes(db: AccessDatabase, tmp_path: Path) -> None:
    db.drop_table("Temp")
    db.compact()
    table = db.table("Keep")
    for number in range(100, 150):
        table.insert_row({"Id": number, "Note": f"after {number}"})
    out = tmp_path / "written.accdb"
    db.save(out)

    assert len(list(AccessDatabase(out).table("Keep").rows())) == 70


def test_compacting_twice_gives_back_nothing_the_second_time(db: AccessDatabase) -> None:
    db.drop_table("Temp")
    assert db.compact() > 0
    assert db.compact() == 0


def test_a_database_with_nothing_to_give_back_is_left_alone(db: AccessDatabase) -> None:
    """Free pages in the middle are not reclaimed: nothing moves."""
    before = db.store.page_count

    assert db.compact() == 0
    assert db.store.page_count == before


def test_a_fresh_database_is_left_alone(tmp_path: Path) -> None:
    db = AccessDatabase.create_new(tmp_path / "blank.accdb")
    before = db.store.page_count

    assert db.compact() == 0
    assert db.store.page_count == before
