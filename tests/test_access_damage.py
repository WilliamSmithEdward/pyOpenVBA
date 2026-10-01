"""A damaged database is refused with an AccessError that says what is wrong.

Issue #45 found Python's own error text reaching users: ``'ParentId'`` from a
KeyError and "unpack requires a buffer of 4 bytes" from struct. Each case
below is a shape the issue or fuzzing produced, reduced to the value that broke.
"""

from __future__ import annotations

import struct
from pathlib import Path

import pytest

from pyopenvba.access import AccessDatabase
from pyopenvba.access._pages import row_bytes
from pyopenvba.access._rows import decode_datetime, decode_scalar
from pyopenvba.access._tdef import _read_name  # pyright: ignore[reportPrivateUsage]
from pyopenvba.access.database import Row
from pyopenvba.access_read import AccessError

BLANK = Path(__file__).parents[1] / "src" / "pyopenvba" / "_templates" / "blank_files" / "blank_database.accdb"


def test_a_column_the_table_lacks_names_the_table_and_column() -> None:
    row = Row("MSysObjects")
    row["Id"] = 1
    assert row["Id"] == 1
    assert row.get("ParentId") is None
    with pytest.raises(AccessError, match="'MSysObjects' has no column 'ParentId'"):
        _ = row["ParentId"]


def test_rows_are_rows() -> None:
    with AccessDatabase(BLANK) as db:
        row = next(db.table("MSysObjects").rows())
    assert isinstance(row, Row)
    with pytest.raises(AccessError, match="has no column 'NoSuchColumn'"):
        _ = row["NoSuchColumn"]


@pytest.mark.parametrize(("column", "raw"), [("Id", b"\x01\x02"), ("DateCreate", b"\x00" * 9)])
def test_a_fixed_size_value_of_the_wrong_length_is_refused(column: str, raw: bytes) -> None:
    with AccessDatabase(BLANK) as db:
        definition = db.table("MSysObjects").definition
    with pytest.raises(AccessError, match=f"column {column!r} holds {len(raw)} bytes"):
        decode_scalar(definition.column(column), raw)


@pytest.mark.parametrize("value", [1e300, float("nan"), float("inf")], ids=["huge", "nan", "inf"])
def test_a_date_no_date_can_hold_is_refused(value: float) -> None:
    with pytest.raises(AccessError, match="outside the range a date can hold"):
        decode_datetime(struct.pack("<d", value))


@pytest.mark.parametrize(
    ("buf", "message"),
    [(b"\x03\x00abc", "odd length"), (b"\x08\x00ab", "past the end"), (b"\x01", "past the end")],
    ids=["odd", "short", "no length"],
)
def test_a_damaged_name_in_a_definition_is_refused(buf: bytes, message: str) -> None:
    with pytest.raises(AccessError, match=message):
        _read_name(buf, 0)


def test_a_slot_past_the_row_table_is_refused() -> None:
    with AccessDatabase(BLANK) as db:
        page = db.store.read(next(iter(db.table("MSysObjects").data_pages())))
    with pytest.raises(AccessError, match="slot 255 out of range"):
        row_bytes(page, 255)
