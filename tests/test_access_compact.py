"""Giving back the free pages a file ends with.

This is not Access's Compact and Repair, which rebuilds the database and
moves every page.  It reclaims the run of free pages at the end, which is
what a dropped table or a large delete leaves behind, and moves nothing.
"""

from __future__ import annotations

from pathlib import Path

import pytest

from pyopenvba.access import AccessDatabase, ColumnSpec, IndexSpec
from pyopenvba.access_read import AccessError
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


def emptied(path: Path, rows: int = 2000, keep: int = 10) -> AccessDatabase:
    """A table most of whose rows have been deleted, saved and reopened --
    which is the state a rebuild is for."""
    db = AccessDatabase.create_new(path)
    table = db.create_table(
        "Wide",
        [
            ColumnSpec("Id", "Long", autonumber=True),
            ColumnSpec("L", "Long", required=True),
            ColumnSpec("T", "Text", size=60, default='"none"'),
            ColumnSpec("M", "Memo"),
            ColumnSpec("O", "OLE"),
        ],
        [IndexSpec("PrimaryKey", ("Id",), primary=True), IndexSpec("ByT", ("T",))],
    )
    for number in range(rows):
        table.insert_row(
            {
                "L": number,
                "T": f"row {number}",
                "M": "memo " * (number % 40 + 1),
                "O": bytes([number % 256]) * (number % 80 + 1),
            }
        )
    db.save(path)
    db = AccessDatabase(path)
    table = db.table("Wide")
    for row_id, row in list(table.rows_with_ids()):
        if int(str(row["L"])) % keep:
            table.delete_row(row_id)
    db.save(path)
    return AccessDatabase(path)


@pytest.fixture
def emptied_db(tmp_path: Path) -> AccessDatabase:
    return emptied(tmp_path / "wide.accdb")


def test_deleting_rows_does_not_shrink_a_table(emptied_db: AccessDatabase) -> None:
    """Which is why a rebuild exists: Access does not give the pages back
    either, and that is what Compact and Repair is for."""
    assert len(list(emptied_db.table("Wide").rows())) == 200
    before = emptied_db.store.page_count
    emptied_db.compact()
    assert emptied_db.store.page_count > before * 0.9, "the pages are still there"


def test_a_rebuild_packs_the_rows_onto_fewer_pages(emptied_db: AccessDatabase) -> None:
    before = emptied_db.store.page_count

    reclaimed = emptied_db.compact(rebuild=True)

    assert reclaimed > 0
    assert emptied_db.store.page_count < before // 2


def test_every_row_survives_a_rebuild(emptied_db: AccessDatabase, tmp_path: Path) -> None:
    before = list(emptied_db.table("Wide").rows())
    emptied_db.compact(rebuild=True)
    out = tmp_path / "packed.accdb"
    emptied_db.save(out)

    assert list(AccessDatabase(out).table("Wide").rows()) == before


def test_the_autonumbers_and_their_counter_are_kept(emptied_db: AccessDatabase) -> None:
    """The rows keep the keys they were given, and the next one carries on
    from where it was -- Access's own compact resets the counter instead,
    which this deliberately does not."""
    ids = [row["Id"] for row in emptied_db.table("Wide").rows()]
    counter = emptied_db.table("Wide").definition.next_autonumber

    emptied_db.compact(rebuild=True)

    table = emptied_db.table("Wide")
    assert [row["Id"] for row in table.rows()] == ids
    assert table.definition.next_autonumber == counter


def test_the_indexes_come_back(emptied_db: AccessDatabase) -> None:
    emptied_db.compact(rebuild=True)

    definition = emptied_db.table("Wide").definition
    assert sorted(i.name for i in definition.logical_indexes) == ["ByT", "PrimaryKey"]
    primary = definition.primary_key()
    assert primary is not None and primary.name == "PrimaryKey"


def test_the_column_rules_come_back(emptied_db: AccessDatabase) -> None:
    emptied_db.compact(rebuild=True)

    table = emptied_db.table("Wide")
    assert table.column_properties("L").get("Required") is True
    assert table.column_properties("T").get("DefaultValue") == '"none"'


def test_the_specs_describe_the_table_as_it_stands(emptied_db: AccessDatabase) -> None:
    columns, indexes = emptied_db.table_specs("Wide")

    assert [c.name for c in columns] == ["Id", "L", "T", "M", "O"]
    text = next(c for c in columns if c.name == "T")
    assert text.size == 60, "a Text size is characters, not the header's bytes"
    assert next(c for c in columns if c.name == "Id").autonumber
    assert sorted(i.name for i in indexes) == ["ByT", "PrimaryKey"]
    assert next(i for i in indexes if i.name == "PrimaryKey").primary


def test_a_table_in_a_relationship_is_refused(tmp_path: Path) -> None:
    """A rebuild drops the table, which would leave the relationship
    naming something that is not there."""
    db = AccessDatabase.create_new(tmp_path / "linked.accdb")
    db.create_table("Parent", [ColumnSpec("Id", "Long")], [IndexSpec("PrimaryKey", ("Id",), primary=True)])
    db.create_table("Child", [ColumnSpec("Id", "Long"), ColumnSpec("ParentId", "Long")],
                    [IndexSpec("PrimaryKey", ("Id",), primary=True)])
    db.create_relationship("ParentChild", "Child", ("ParentId",), "Parent", ("Id",))

    with pytest.raises(AccessError, match="relationship"):
        db.rebuild_table("Child")


def test_a_refusal_leaves_the_database_alone(tmp_path: Path) -> None:
    db = AccessDatabase.create_new(tmp_path / "linked.accdb")
    db.create_table("Parent", [ColumnSpec("Id", "Long")], [IndexSpec("PrimaryKey", ("Id",), primary=True)])
    db.create_table("Child", [ColumnSpec("Id", "Long"), ColumnSpec("ParentId", "Long")],
                    [IndexSpec("PrimaryKey", ("Id",), primary=True)])
    db.create_relationship("ParentChild", "Child", ("ParentId",), "Parent", ("Id",))
    kept = db.to_bytes()

    with pytest.raises(AccessError):
        db.rebuild_table("Child")

    assert db.to_bytes() == kept


def test_compacting_with_rebuild_leaves_a_refused_table_alone(tmp_path: Path) -> None:
    """One table it will not touch does not stop the rest of the file."""
    db = emptied(tmp_path / "mixed.accdb")
    db.create_table("Parent", [ColumnSpec("Id", "Long")], [IndexSpec("PK2", ("Id",), primary=True)])
    db.create_table("Child", [ColumnSpec("Id", "Long"), ColumnSpec("ParentId", "Long")],
                    [IndexSpec("PK3", ("Id",), primary=True)])
    db.create_relationship("ParentChild", "Child", ("ParentId",), "Parent", ("Id",))
    held = len(db.table("Wide").data_pages())

    db.compact(rebuild=True)

    # The pair in the relationship is left as it was, and the table that
    # could be packed was: the trailing run is another matter, since the
    # pair was created last and sits at the end of the file.
    assert len(db.table("Wide").data_pages()) < held // 2
    assert sorted(db.table_names()) == ["Child", "Parent", "Wide"]
    assert len(list(db.table("Wide").rows())) == 200
    assert "ParentChild" in [r.name for r in db.relationships()]
