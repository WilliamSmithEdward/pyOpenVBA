"""Compaction: giving back the free pages a file ends with, and Compact and
Repair proper.

`compact()` reclaims the run of free pages at the end, which is what a
dropped table or a large delete leaves behind, and moves nothing.
`compact_and_repair()` rebuilds the database into a fresh engine skeleton
the way DAO's CompactDatabase does; the live gate holds it to the engine's
bytes, these tests to what it must carry across.
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


# -- Compact and Repair --------------------------------------------------------


def related(path: Path) -> AccessDatabase:
    """A database with what a compaction has to carry across: an AutoNumber
    table with a memo column and deleted rows, a keyed table written out
    of key order, a heap AutoNumber table with explicit ids, a saved
    query and a relationship."""
    db = AccessDatabase.create_new(path)
    orders = db.create_table(
        "Orders",
        [ColumnSpec("Id", "Long", autonumber=True), ColumnSpec("Name", "Text", size=30), ColumnSpec("Notes", "Memo")],
        [IndexSpec("PrimaryKey", ("Id",), primary=True), IndexSpec("ByName", ("Name",))],
    )
    for number in range(1, 41):
        orders.insert_row({"Name": f"order {number}", "Notes": "memo " * (number % 30) or None})
    for row_id, row in list(orders.rows_with_ids()):
        if int(str(row["Id"])) % 3 == 0 or int(str(row["Id"])) > 35:
            orders.delete_row(row_id)
    lines = db.create_table(
        "Lines",
        [ColumnSpec("Id", "Long"), ColumnSpec("OrderId", "Long"), ColumnSpec("Qty", "Long")],
        [IndexSpec("PrimaryKey", ("Id",), primary=True)],
    )
    for ident, order in ((5, 1), (3, 2), (9, 4), (1, 5)):
        lines.insert_row({"Id": ident, "OrderId": order, "Qty": ident * 2})
    heap = db.create_table("Heap", [ColumnSpec("Id", "Long", autonumber=True), ColumnSpec("V", "Text", size=5)])
    for ident in (5, 3, 9, 1):
        heap.insert_row({"Id": ident, "V": str(ident)})
    db.create_query("QOrders", "SELECT Orders.Id, Orders.Name FROM Orders")
    db.create_relationship("FK_Lines_Orders", "Lines", ("OrderId",), "Orders", ("Id",))
    db.save(path)
    return AccessDatabase(path)


@pytest.fixture
def related_db(tmp_path: Path) -> AccessDatabase:
    return related(tmp_path / "related.accdb")


def test_the_skeleton_is_the_bare_engine_database() -> None:
    """What the copy starts from: DAO's CreateDatabase output with its
    permission rows taken back out."""
    from pyopenvba.access._compact import SKELETON

    skeleton = AccessDatabase(SKELETON.read_bytes())
    assert skeleton.store.page_count == 41
    assert skeleton.table("MSysACEs").row_count == 0
    names = {e.name for e in skeleton.catalog()}
    assert {"Tables", "Databases", "Relationships", "MSysDb", "MSysObjects", "MSysACEs", "MSysQueries", "MSysRelationships", "MSysComplexColumns"} <= names
    assert len(skeleton.catalog()) == 18
    from pyopenvba.access._compact import SKELETON_JET4

    jet4 = AccessDatabase(SKELETON_JET4.read_bytes())
    assert (jet4.store.page_count, jet4.table("MSysACEs").row_count, len(jet4.catalog())) == (18, 0, 8)


def test_compact_and_repair_carries_every_object_and_row(related_db: AccessDatabase, tmp_path: Path) -> None:
    before = {name: sorted(related_db.table(name).rows(), key=lambda r: int(str(r["Id"]))) for name in ("Orders", "Lines", "Heap")}
    compacted = related_db.compact_and_repair()
    assert compacted.store.page_count < related_db.store.page_count
    assert set(compacted.table_names()) == set(related_db.table_names())
    assert [q.name for q in compacted.queries()] == ["QOrders"]
    # Relationships come back in Name order, whatever order they were made in.
    assert [r.name for r in compacted.relationships()] == sorted((r.name for r in related_db.relationships()), key=str.lower)
    for name, rows in before.items():
        assert sorted(compacted.table(name).rows(), key=lambda r: int(str(r["Id"]))) == rows
    # A keyed table's rows come out in key order; a heap keeps its order.
    assert [r["Id"] for r in compacted.table("Lines").rows()] == [1, 3, 5, 9]
    assert [r["Id"] for r in compacted.table("Heap").rows()] == [5, 3, 9, 1]
    # The catalog and permission rows all come across, and the file reopens.
    assert len(compacted.catalog()) == len(related_db.catalog())
    assert compacted.table("MSysACEs").row_count == related_db.table("MSysACEs").row_count
    out = tmp_path / "compacted.accdb"
    compacted.save(out)
    assert AccessDatabase(out).table("Orders").row_count == related_db.table("Orders").row_count


def test_compact_and_repair_resets_each_autonumber_to_the_largest_value_present(related_db: AccessDatabase) -> None:
    assert related_db.table("Orders").definition.next_autonumber == 40
    compacted = related_db.compact_and_repair()
    assert compacted.table("Orders").definition.next_autonumber == 35
    assert compacted.table("Heap").definition.next_autonumber == 9
    # And the next row takes the number after it.
    compacted.table("Orders").insert_row({"Name": "next"})
    assert max(int(str(r["Id"])) for r in compacted.table("Orders").rows()) == 36


def test_compact_and_repair_leaves_the_source_alone(related_db: AccessDatabase) -> None:
    image = related_db.to_bytes()
    related_db.compact_and_repair()
    assert related_db.to_bytes() == image


def test_compact_and_repair_keeps_the_creation_date_and_the_owners(related_db: AccessDatabase) -> None:
    """The engine keys the SID encoding to the creation date; keeping the
    date keeps every owner and permission SID as it was."""
    compacted = related_db.compact_and_repair()
    assert compacted.header.creation_date == related_db.header.creation_date
    owners = {e.name: e.owner for e in related_db.catalog()}
    assert {e.name: e.owner for e in compacted.catalog()} == owners


# -- a wide table's stale definition pages -------------------------------------
#
# The engine builds a table whose definition runs past one page column by
# column, two writes per column onto fresh continuation pages, and lets the
# pages it gives back come into use again only as the destination file
# grows (64 pages ahead, then 8 or 64 at a time, the waiting pages returning
# once 24 pages have been taken since they last did).  These layouts are the
# engine's, read off DAO's compactions of the same tables; the labels are
# page:kind, a definition page's kind being the column count its bytes hold.


#: The template Access starts a database from; ``create_new`` uses a
#: smaller one, and the engine's layouts below were read off compactions
#: of databases made from this one.
_TEMPLATE = Path(__file__).resolve().parents[1] / "src" / "pyopenvba" / "_templates" / "blank_files" / "blank_database.accdb"


def _wide_columns(count: int, memo: bool = False) -> list[ColumnSpec]:
    columns = [ColumnSpec("Id", "Long")]
    if memo:
        columns.append(ColumnSpec("M", "Memo"))
    return columns + [ColumnSpec(f"C{i:03}", "Text", size=20) for i in range(1, count + 1)]


def _layout(db: AccessDatabase, table: str) -> str:
    """Every page from the table's definition page on: ``T`` its first
    page, ``M`` a usage-map page, ``cN`` a definition page holding N
    columns (``L`` when live), ``D<owner>`` a data page, ``I<owner>`` an
    index page, ``LV`` a long-value page."""
    import struct

    store = db.store
    definition = db.table(table).definition
    out: list[str] = []
    for page in range(definition.page, store.page_count):
        raw = store.read(page)
        kind = raw[0]
        if kind == 2:
            if page == definition.page:
                out.append(f"{page}:T")
            else:
                columns = (2 * (PAGE_SIZE - 8) - struct.unpack_from("<H", raw, 2)[0] - 96) / 35
                out.append(f"{page}:c{columns:.0f}{'L' if page in definition.pages else ''}")
        elif kind == 1:
            owner = struct.unpack_from("<I", raw, 4)[0]
            out.append(f"{page}:{'LV' if raw[4:8] == b'LVAL' else 'M' if owner == 0 else 'D' + str(owner)}")
        else:
            out.append(f"{page}:I{struct.unpack_from('<I', raw, 4)[0]}")
    return " ".join(out)


def test_a_wide_table_leaves_the_stale_definition_pages_the_engine_leaves() -> None:
    """131 columns: the definition runs past one page from the 116th, and
    the copy cycles through 17 continuation pages, the release coming
    when the file reaches 128 pages."""
    db = AccessDatabase(_TEMPLATE.read_bytes())
    db.create_table("Wide", _wide_columns(130))

    compacted = db.compact_and_repair()

    assert _layout(compacted, "Wide") == (
        "109:T 110:M 111:c123 112:c124 113:c124 114:c125 115:c125 116:c126 117:c126 118:c127 119:c127 "
        "120:c128 121:c128 122:c129 123:c129 124:c130 125:c130 126:c130L 127:c123 128:D18 129:D3 130:LV 131:D3 132:D5"
    )


def test_a_keyed_wide_table_with_rows_is_rewritten_after_the_rows_and_twice_for_the_index() -> None:
    """161 columns, a primary key and two rows: the data page comes before
    the write after the rows, the index root before the two writes with
    the index, and the pages given back come round three times at 136."""
    db = AccessDatabase(_TEMPLATE.read_bytes())
    table = db.create_table("Wide", _wide_columns(160), [IndexSpec("PK", ("Id",), primary=True)])
    table.insert_row({"Id": 1, "C001": "one"})
    table.insert_row({"Id": 2, "C001": "two"})

    compacted = db.compact_and_repair()

    assert _layout(compacted, "Wide") == (
        "109:T 110:M 111:c159 112:c160 113:c160 114:D109 115:c160 116:I109 117:c163 118:c163L 119:D18 120:D3 121:LV 122:D3 123:D5 "
        "124:c154 125:c154 126:c155 127:c155 128:c156 129:c156 130:c157 131:c157 132:c158 133:c158 134:c147 135:c159"
    )


def test_two_wide_tables_grow_the_file_by_the_engine_s_steps() -> None:
    """A wide table copied first cycles within 16 pages and grows the file
    from 64 to 72; the system tables then take those pages back, grow it
    to 136, and the second wide table's pages come back at 136 and 144."""
    db = AccessDatabase(_TEMPLATE.read_bytes())
    db.create_table("Bide", _wide_columns(160))
    db.create_table("Wide", _wide_columns(160))

    compacted = db.compact_and_repair()

    assert compacted.table("Bide").definition.pages == [41, 58]
    assert _layout(compacted, "Wide") == (
        "112:T 113:M 114:c155 115:c155 116:c156 117:c156 118:c157 119:c157 120:c158 121:c158 122:c159 123:c159 124:c160 125:c160 126:c160L "
        "127:D18 128:D3 129:LV 130:D3 131:D2 132:D5 133:c150 134:c150 135:c151 136:c151 137:c152 138:c152 139:c153 140:c153 141:c154 142:c154 143:c140"
    )
