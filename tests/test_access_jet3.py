"""Reading Jet 3, the Access 97 format.

Jet 3 halves the page, counts a row's columns in a byte rather than a
word, stores text in the database code page rather than UTF-16, and moves
almost every field of a table definition.  All of that lives in
``_layout.Layout``, so the same parser, row splitter and value decoders
read both versions.

The fixture was written by the Jet engine itself through DAO 3.6, which
still creates Access 97 databases even though Access dropped the format
in 2013.  ``test_live_access_jet3_gate.py`` checks a larger one against
the engine field for field.

Writing a Jet 3 file is a different engine, not a different set of
offsets, so it is refused rather than attempted.
"""

from __future__ import annotations

import datetime as dt
from decimal import Decimal
from pathlib import Path

import pytest

from pyopenvba.access import AccessDatabase
from pyopenvba.access._layout import JET3, JET4
from pyopenvba.access._pages import PageStore
from pyopenvba.exceptions import UnsupportedFormatError

FIXTURE = Path(__file__).parent / "live_access_test" / "jet3_orders.mdb"


@pytest.fixture
def db() -> AccessDatabase:
    return AccessDatabase(FIXTURE)


def test_the_file_says_which_version_it_is(db: AccessDatabase) -> None:
    assert db.store.is_jet3
    assert db.store.layout is JET3
    assert db.store.page_size == 2048
    assert db.store.code_page == "cp1252"


def test_an_accdb_still_reads_as_jet4(tmp_path: Path) -> None:
    fresh = AccessDatabase.create_new(tmp_path / "new.accdb")
    assert not fresh.store.is_jet3
    assert fresh.store.layout is JET4
    assert fresh.store.page_size == 4096


def test_the_catalog_names_the_tables(db: AccessDatabase) -> None:
    assert db.table_names() == ["Lines", "Orders"]
    assert "MSysObjects" in db.table_names(include_system=True)


def test_the_columns_come_back_with_their_types(db: AccessDatabase) -> None:
    table = db.table("Orders")
    assert table.column_names == ["Id", "Customer", "Total", "Placed", "Comment"]
    types = {c.name: c.type_name for c in table.definition.columns}
    assert types == {
        "Id": "Long",
        "Customer": "Text",
        "Total": "Currency",
        "Placed": "DateTime",
        "Comment": "Memo",
    }


def test_every_value_decodes(db: AccessDatabase) -> None:
    rows = list(db.table("Orders").rows())
    assert rows[0] == {
        "Id": 1,
        "Customer": "Ada",
        "Total": Decimal("10.5000"),
        "Placed": dt.datetime(2001, 2, 3, 4, 5, 6),
        "Comment": "first",
    }
    assert rows[2]["Total"] is None and rows[2]["Comment"] is None


def test_text_comes_back_in_the_code_page_not_as_utf16(db: AccessDatabase) -> None:
    """Jet 3 stores one byte per character; read as UTF-16 the name would
    come back as a single CJK character instead of four Latin ones."""
    rows = list(db.table("Orders").rows())
    assert rows[1]["Customer"] == "Bob\u00e9"


def test_a_memo_too_long_for_its_row_is_followed(db: AccessDatabase) -> None:
    """1500 characters do not fit a 2 KiB page, so the value is a chain of
    long-value rows."""
    rows = list(db.table("Orders").rows())
    comment = rows[1]["Comment"]
    assert isinstance(comment, str)
    assert len(comment) == 1500
    assert comment.startswith("wide wide") and comment.endswith("wide ")


def test_a_deleted_row_is_not_returned(db: AccessDatabase) -> None:
    assert [r["Customer"] for r in db.table("Orders").rows()] == [
        "Ada",
        "Bob\u00e9",
        "no totals",
    ]


def test_rows_spanning_several_pages_all_come_back(db: AccessDatabase) -> None:
    rows = sorted(db.table("Lines").rows(), key=lambda r: int(str(r["Id"])))
    assert len(rows) == 60
    assert rows[0]["Id"] == 1 and rows[0]["OrderId"] == 2 and rows[0]["Qty"] == 2
    assert rows[-1]["Id"] == 60 and rows[-1]["OrderId"] == 1 and rows[-1]["Qty"] == 120
    assert rows[-1]["Item"] == "item 60 " + "pad " * 20
    assert len(db.table("Lines").data_pages()) > 1


def test_the_primary_key_is_recognised(db: AccessDatabase) -> None:
    primary = db.table("Orders").definition.primary_key()
    assert primary is not None and primary.name == "PK"


def test_a_definition_accounts_for_every_byte_it_declares(db: AccessDatabase) -> None:
    """The parser checks that what it consumed equals the length the page
    declares, so this passing is what says the Jet 3 offsets are right."""
    for name in db.table_names(include_system=True):
        definition = db.table(name).definition
        assert definition.definition_length > 0
        assert definition.layout is JET3


def test_writing_is_refused_rather_than_attempted(db: AccessDatabase) -> None:
    with pytest.raises(UnsupportedFormatError, match="read-only"):
        db.table("Orders").insert_row({"Customer": "nope"})
    with pytest.raises(UnsupportedFormatError, match="read-only"):
        db.drop_table("Lines")
    with pytest.raises(UnsupportedFormatError, match="read-only"):
        db.create_module("M", "Sub S()\nEnd Sub")


def test_a_page_cannot_be_written_even_from_underneath(db: AccessDatabase) -> None:
    """The guard that matters is the lowest one: nothing can put a Jet 4
    page into a Jet 3 file."""
    store = PageStore(FIXTURE.read_bytes())
    with pytest.raises(UnsupportedFormatError, match="read-only"):
        store.write(3, bytes(store.page_size))


def test_the_file_can_still_be_copied_out(db: AccessDatabase, tmp_path: Path) -> None:
    out = tmp_path / "copy.mdb"
    db.save(out)
    assert out.read_bytes() == FIXTURE.read_bytes()
    assert len(list(AccessDatabase(out).table("Lines").rows())) == 60
