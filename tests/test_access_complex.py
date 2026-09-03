"""Attachment and multi-valued columns.

The fixture was built by DAO: `Things` has an attachment column `Files`
and a multi-valued text column `Tags`, three rows, and one row with
neither so an empty complex value is covered.  The live gate hands what
this writes back to the engine and reads it through DAO.
"""

from __future__ import annotations

import shutil
import zlib
from pathlib import Path

import pytest

from pyopenvba.access import AccessDatabase, Attachment
from pyopenvba.access._complex import (
    NEVER_COMPRESSED,
    decode_file_data,
    encode_file_data,
)
from pyopenvba.access_read import AccessError

FIXTURE = Path(__file__).parent / "live_access_test" / "complex_columns.accdb"


@pytest.fixture
def db(tmp_path: Path) -> AccessDatabase:
    out = tmp_path / "complex.accdb"
    shutil.copyfile(FIXTURE, out)
    return AccessDatabase(out)


def key_for(db: AccessDatabase, name: str, column: str) -> int:
    row = next(r for r in db.table("Things").rows() if r["Name"] == name)
    return int(str(row[column]))


# --- the FileData container ---------------------------------------------------


def test_every_attachment_in_the_fixture_decodes(db: AccessDatabase) -> None:
    things = db.table("Things")
    assert [a.name for a in things.attachments("Files", key_for(db, "first", "Files"))] == [
        "attach_one.txt",
        "attach_two.txt",
    ]
    first = things.attachments("Files", key_for(db, "first", "Files"))[0]
    assert first.data == b"hello from one\r\n"
    assert first.type == "txt"


@pytest.mark.parametrize(
    ("extension", "data"),
    [
        ("txt", b"hello from one\r\n"),
        ("7z", b"A" * 512),  # a two-character extension: an 18-byte header
        ("docx", b"A" * 512),  # four characters, and never compressed
        ("png", bytes(range(64))),
        ("", b"no extension at all"),
    ],
)
def test_file_data_round_trips(extension: str, data: bytes) -> None:
    blob = encode_file_data(extension, data)
    assert decode_file_data(blob) == (extension, data)


def test_the_inner_header_grows_with_the_extension() -> None:
    """12 fixed bytes plus the NUL-terminated extension in UTF-16, which
    is 20 for `txt` and 22 for `docx`."""
    for extension, expected in (("7z", 18), ("txt", 20), ("docx", 22)):
        blob = encode_file_data(extension, b"")
        body = zlib.decompress(blob[8:]) if blob[0] == 1 else blob[8:]
        assert int.from_bytes(body[0:4], "little") == expected


def test_the_types_access_leaves_alone_are_stored_raw() -> None:
    for extension in sorted(NEVER_COMPRESSED):
        blob = encode_file_data(extension, b"A" * 512)
        assert int.from_bytes(blob[0:4], "little") == 0, extension
    assert int.from_bytes(encode_file_data("txt", b"A" * 512)[0:4], "little") == 1


def test_a_stored_raw_attachment_is_byte_identical_to_access(db: AccessDatabase) -> None:
    """Compression is the one thing that cannot match -- Access's deflate
    is not zlib's -- so the check is on a type it does not compress."""
    original = encode_file_data("png", bytes(range(64)))
    extension, data = decode_file_data(original)
    assert encode_file_data(extension, data) == original


def test_a_short_or_unknown_container_is_refused() -> None:
    with pytest.raises(AccessError, match="too short"):
        decode_file_data(b"\x00\x00\x00")
    with pytest.raises(AccessError, match="unknown storage flag"):
        decode_file_data((7).to_bytes(4, "little") + (0).to_bytes(4, "little"))


# --- discovery ----------------------------------------------------------------


def test_complex_columns_name_their_flat_tables(db: AccessDatabase) -> None:
    found = {c.column: c for c in db.complex_columns("Things")}
    assert set(found) == {"Files", "Tags"}
    assert found["Files"].kind == "attachment" and found["Files"].is_attachment
    assert found["Tags"].kind == "Text" and not found["Tags"].is_attachment
    assert found["Files"].flat_table.startswith("f_") and found["Files"].flat_table.endswith("_Files")
    assert found["Files"].key_column == "_Files"
    assert found["Files"].id_column == "Things_Files"
    assert db.table("Things").complex_columns() == db.complex_columns("Things")


def test_reading_covers_an_empty_complex_value(db: AccessDatabase) -> None:
    things = db.table("Things")
    assert things.attachments("Files", key_for(db, "third", "Files")) == []
    assert things.multi_values("Tags", key_for(db, "third", "Tags")) == []


def test_elements_come_back_in_the_order_the_engine_walks_them(db: AccessDatabase) -> None:
    """The flat table's `(key, Value)` index decides, so values come back
    sorted rather than in the order they were added -- which is what DAO
    hands back for the same rows."""
    things = db.table("Things")
    assert things.multi_values("Tags", key_for(db, "first", "Tags")) == ["alpha", "beta", "gamma"]
    assert things.multi_values("Tags", key_for(db, "second", "Tags")) == ["delta"]
    things.set_multi_values("Tags", key_for(db, "second", "Tags"), ["red", "green", "blue"])
    assert things.multi_values("Tags", key_for(db, "second", "Tags")) == ["blue", "green", "red"]


def test_the_wrong_kind_of_column_is_refused(db: AccessDatabase) -> None:
    things = db.table("Things")
    with pytest.raises(AccessError, match="not attachments"):
        things.attachments("Tags", 1)
    with pytest.raises(AccessError, match="not scalar values"):
        things.multi_values("Files", 1)
    with pytest.raises(AccessError, match="no complex column named"):
        things.attachments("Name", 1)


# --- writing ------------------------------------------------------------------


def test_attachments_can_be_replaced(db: AccessDatabase, tmp_path: Path) -> None:
    things = db.table("Things")
    key = key_for(db, "third", "Files")
    things.set_attachments("Files", key, [Attachment("readme.txt", b"written here")])
    out = tmp_path / "written.accdb"
    db.save(out)

    reopened = AccessDatabase(out).table("Things")
    files = reopened.attachments("Files", key)
    assert [(a.name, a.data, a.type) for a in files] == [("readme.txt", b"written here", "txt")]


def test_multi_values_can_be_replaced(db: AccessDatabase, tmp_path: Path) -> None:
    things = db.table("Things")
    key = key_for(db, "first", "Tags")
    things.set_multi_values("Tags", key, ["one", "two"])
    out = tmp_path / "written.accdb"
    db.save(out)

    assert AccessDatabase(out).table("Things").multi_values("Tags", key) == ["one", "two"]


def test_an_inserted_row_takes_the_next_complex_id(db: AccessDatabase) -> None:
    """Every complex column in a row shares one id, from the counter at
    0x1C, and a row with no elements still takes one."""
    things = db.table("Things")
    before = things.definition.last_complex_id
    assert before == 3

    things.insert_row({"Id": 7, "Name": "seventh"})
    row = next(r for r in things.rows() if r["Id"] == 7)
    assert row["Files"] == before + 1
    assert row["Tags"] == before + 1
    assert things.definition.last_complex_id == before + 1

    things.insert_row({"Id": 8, "Name": "eighth"})
    assert things.definition.last_complex_id == before + 2


def test_a_row_written_end_to_end_reads_back(db: AccessDatabase, tmp_path: Path) -> None:
    things = db.table("Things")
    things.insert_row({"Id": 7, "Name": "seventh"})
    row = next(r for r in things.rows() if r["Id"] == 7)
    key = int(str(row["Files"]))
    things.set_attachments(
        "Files", key, [Attachment("a.txt", b"first file"), Attachment("b.png", bytes(range(64)))]
    )
    things.set_multi_values("Tags", int(str(row["Tags"])), ["red", "green", "blue"])
    out = tmp_path / "written.accdb"
    db.save(out)

    reopened = AccessDatabase(out).table("Things")
    assert [(a.name, len(a.data)) for a in reopened.attachments("Files", key)] == [
        ("a.txt", 10),
        ("b.png", 64),
    ]
    assert reopened.multi_values("Tags", key) == ["blue", "green", "red"]


def test_an_attachment_takes_its_type_from_its_name() -> None:
    assert Attachment("report.PDF", b"").type == "pdf"
    assert Attachment("no_extension", b"").type == ""
    assert Attachment("given.txt", b"", type="dat").type == "dat"
