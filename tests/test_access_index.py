"""B-tree index reading, against the Access-authored fixtures.

The invariants checked are the ones a wrong page layout or key codec
would break: every index has as many leaf entries as the table has rows
(nulls aside when the index ignores them), entries come out in
non-decreasing key order, every entry points at a live row whose decoded
key equals the row's own values, and every node entry names the last key
of its child.
"""

from __future__ import annotations

import decimal
import shutil
from pathlib import Path

import pytest

from pyopenvba.access import AccessDatabase
from pyopenvba.access._index import (
    TextKey,
    decode_key,
    encode_key,
    leaf_entries,
    node_pages,
    parse_index_page,
)
from pyopenvba.access._rows import decode_numeric, encode_numeric, split_row
from pyopenvba.access._tdef import TYPE_NUMERIC, ColumnDef

FIXTURES = Path(__file__).parent / "live_access_test"
TEMPLATES = Path(__file__).parents[1] / "src" / "pyopenvba" / "_templates" / "blank_files"
AUTHORED = [
    FIXTURES / "New Microsoft Access Database.accdb",
    FIXTURES / "module_spanning_pages.accdb",
    FIXTURES / "two_modules_one_page.accdb",
    TEMPLATES / "blank_database.accdb",
    TEMPLATES / "blank_database_module.accdb",
]
SMALL = FIXTURES / "two_modules_one_page.accdb"
LARGE = FIXTURES / "New Microsoft Access Database.accdb"


def _rows_by_home(table: "object") -> dict[tuple[int, int], dict[str, object]]:
    from pyopenvba.access.database import Table

    assert isinstance(table, Table)
    out: dict[tuple[int, int], dict[str, object]] = {}
    for page, slot, data in table.raw_rows():
        out[(page, slot)] = table.decode(split_row(table.definition, data))
    return out


@pytest.mark.parametrize("path", AUTHORED, ids=lambda p: p.name)
def test_every_index_agrees_with_its_table(path: Path) -> None:
    db = AccessDatabase(path)
    checked = 0
    for table in db.tables(include_system=True):
        rows = _rows_by_home(table)
        for index in table.indexes:
            columns = index.columns
            expected = len(rows)
            if index.ignores_nulls:
                expected -= sum(
                    1 for r in rows.values() if all(r[c.name] is None for c, _ in columns)
                )
            previous: bytes | None = None
            count = 0
            for entry in leaf_entries(db.store, index.real.root_page):
                count += 1
                assert previous is None or entry.key >= previous, (table.name, index.name)
                previous = entry.key
                assert (entry.page, entry.row) in rows, (table.name, index.name, entry)
                row = rows[(entry.page, entry.row)]
                for (column, _asc), value in zip(columns, decode_key(entry.key, columns)):
                    _assert_key_matches(column, value, row[column.name])
            assert count == expected, (table.name, index.name, count, expected)
            for node in node_pages(db.store, index.real.root_page):
                for node_entry in node.entries:
                    assert node_entry.child is not None
                    child = parse_index_page(db.store, node_entry.child)
                    last = child.entries[-1]
                    assert (last.key, last.page, last.row) == (
                        node_entry.key, node_entry.page, node_entry.row,
                    ), (table.name, index.name)
            checked += 1
    assert checked >= 20


def _assert_key_matches(column: ColumnDef, key_value: object, row_value: object) -> None:
    if key_value is None:
        # A null key for a Boolean never happens; for any other type the
        # row must be null too.
        assert row_value is None or column.type_code == 0x01
        return
    if isinstance(key_value, TextKey):
        assert isinstance(row_value, str)
        # An empty primary means every character was ignorable; otherwise
        # the key must carry at least one byte per significant character.
        assert key_value.primary or not row_value.strip(" ")
        return
    if isinstance(key_value, float) and isinstance(row_value, float):
        assert abs(key_value - row_value) <= 1e-9 * max(1.0, abs(row_value))
        return
    assert key_value == row_value, (column.name, key_value, row_value)


def test_msysobjects_indexes_are_named_and_typed() -> None:
    db = AccessDatabase(SMALL)
    table = db.table("MSysObjects")
    names = {i.name: i for i in table.indexes}
    assert set(names) == {"Id", "ParentIdName"}
    pk = table.primary_key
    assert pk is not None and pk.name == "Id" and pk.unique
    assert names["ParentIdName"].column_names == ["ParentId", "Name"]
    assert names["ParentIdName"].unique
    assert names["ParentIdName"].distinct_count == table.row_count


def test_rows_come_back_in_key_order() -> None:
    db = AccessDatabase(SMALL)
    table = db.table("MSysObjects")
    assert table.primary_key is not None
    ids: list[int] = []
    for row in table.primary_key.rows():
        assert isinstance(row["Id"], int)
        ids.append(row["Id"])
    assert ids == sorted(ids)
    assert len(ids) == table.row_count
    parents: list[int] = []
    for key, _page, _row in table.index("ParentIdName").entries():
        assert isinstance(key[0], int)
        parents.append(key[0])
    assert parents == sorted(parents)
    assert len(parents) == table.row_count


def test_text_keys_are_case_insensitive_collation_bytes() -> None:
    """'Modules' and 'MSysObjects' share their first letter: the collation
    byte for M is the same whether the row spells it upper or lower case,
    and the key holds no case information at all."""
    db = AccessDatabase(SMALL)
    keys: dict[str, TextKey] = {}
    table = db.table("MSysObjects")
    rows = _rows_by_home(table)
    index = table.index("ParentIdName")
    for values, page, row in index.entries():
        name = rows[(page, row)]["Name"]
        text = values[1]
        if isinstance(text, TextKey) and isinstance(name, str):
            keys[name] = text
    assert keys["Modules"].primary[0] == keys["MSysObjects"].primary[0] == 0x60
    assert keys["Tables"].primary[0] == 0x6D
    assert all(k.extra == b"" for k in keys.values())


def test_overflow_row_is_reachable_through_its_index() -> None:
    db = AccessDatabase(LARGE)
    table = db.table("MSysObjects")
    names = {r["Name"] for r in table.index("ParentIdName").rows()}
    assert "Table2" in names


# --- Decimal keys and the arithmetic context ---------------------------------
# A Decimal column is a 16-byte magnitude, so a value carries up to 39
# digits, while Python's default arithmetic context rounds at 28.  Both
# codecs scaled through that context, so the top of the range moved and
# the key stopped naming the row it pointed at (GitHub issue #20).

#: 2**96 - 1, the largest magnitude the storage holds, and 29 digits.
DECIMAL_MAX = decimal.Decimal("79228162514264337593543950335")


def test_the_default_context_is_the_one_that_rounds() -> None:
    """The premise, pinned: this is why the codecs cannot use bare
    arithmetic.  If Python's default precision ever changes, the fix
    below stops being load-bearing and this says so."""
    assert decimal.getcontext().prec == 28
    assert int(DECIMAL_MAX.scaleb(0)) != int(DECIMAL_MAX)
    assert int(DECIMAL_MAX.scaleb(0)) == int(DECIMAL_MAX) + 5


def test_the_row_codec_carries_the_whole_range() -> None:
    for value in (DECIMAL_MAX, -DECIMAL_MAX, decimal.Decimal(0)):
        assert decode_numeric(encode_numeric(value, 0), 0) == value


def test_a_scaled_decimal_keeps_every_digit() -> None:
    value = decimal.Decimal("1234567890123456789012345.6789")
    assert decode_numeric(encode_numeric(value, 4), 4) == value


def _decimal_table(tmp_path: Path) -> Path:
    """A one-row table whose Decimal column holds the largest magnitude
    the storage carries, with an index over it."""
    from pyopenvba.access import ColumnSpec, IndexSpec

    out = tmp_path / "decimal.accdb"
    shutil.copyfile(TEMPLATES / "blank_database.accdb", out)
    database = AccessDatabase(out)
    database.create_table(
        "D", [ColumnSpec("Id", "Long"), ColumnSpec("V", "Decimal", size=(29, 0))]
    )
    database.create_index("D", IndexSpec("ixV", ("V",)))
    database.table("D").insert_row({"Id": 1, "V": DECIMAL_MAX})
    database.save()
    return out


def test_the_index_key_carries_the_whole_range(tmp_path: Path) -> None:
    """The key is the magnitude big-endian behind a sign byte, so a value
    the context rounded shows up in the bytes themselves."""
    table = AccessDatabase(_decimal_table(tmp_path)).table("D")
    column = next(c for c in table.definition.columns if c.name == "V")
    assert column.type_code == TYPE_NUMERIC

    key = encode_key([DECIMAL_MAX], [(column, True)])
    assert key == bytes([0x7F, 0xFF]) + (2**96 - 1).to_bytes(16, "big")
    assert decode_key(key, [(column, True)]) == [DECIMAL_MAX]


def test_a_decimal_row_and_its_index_entry_agree(tmp_path: Path) -> None:
    """End to end, which is what the bug actually broke: an index entry
    that does not match the row it points at sorts wrong and never
    matches on lookup."""
    table = AccessDatabase(_decimal_table(tmp_path)).table("D")

    assert list(table.rows())[0]["V"] == DECIMAL_MAX
    assert [key for key, _page, _row in table.index("ixV").entries()] == [[DECIMAL_MAX]]
    assert [row["V"] for row in table.index("ixV").rows()] == [DECIMAL_MAX]
