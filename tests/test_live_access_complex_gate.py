"""Live ACE gate for attachment and multi-valued columns (opt-in).

A complex value is spread over three places -- the Long in the row, the
flat table's rows, and the container inside `FileData` -- so reading back
what we wrote proves only that we agree with ourselves.  This gate hands
the file to the ACE engine through DAO and compares the bytes DAO saves
out of each attachment with the bytes that went in.

Opt-in: set ``RUN_LIVE_ACCESS=1`` on Windows with the Access database
engine installed.  DAO is a test-time oracle only; pyOpenVBA never uses
COM.
"""

from __future__ import annotations

import json
import os
import shutil
import subprocess
import sys
from collections.abc import Mapping
from dataclasses import dataclass
from pathlib import Path
from typing import cast

import pytest

from pyopenvba.access._complex import encode_file_data
from pyopenvba.access import AccessDatabase, Attachment, ColumnSpec, IndexSpec

pytestmark = pytest.mark.skipif(
    os.environ.get("RUN_LIVE_ACCESS") != "1" or sys.platform != "win32",
    reason="live ACE gate: set RUN_LIVE_ACCESS=1 on Windows with Access installed",
)

HERE = Path(__file__).parent
ORACLE = HERE / "live_access_test" / "dao_complex.ps1"
FIXTURE = HERE / "live_access_test" / "complex_columns.accdb"
_TIMEOUT = 300


@dataclass(frozen=True)
class DaoRow:
    """One row as DAO reports it."""

    id: int
    files: list[tuple[str, str, bytes]]
    tags: list[str]


def through_dao(path: Path, table: str, attachments: str, multi: str) -> dict[int, DaoRow]:
    """What DAO reports for every row of `table`, keyed by Id."""
    done = subprocess.run(
        [
            "powershell",
            "-NoProfile",
            "-NonInteractive",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            str(ORACLE),
            "-Database",
            str(path.resolve()),
            "-Table",
            table,
            "-AttachmentColumn",
            attachments,
            "-MultiValueColumn",
            multi,
        ],
        capture_output=True,
        text=True,
        timeout=_TIMEOUT,
        check=False,
    )
    assert done.returncode == 0, f"DAO oracle failed: {done.stderr.strip()}"
    # A single row comes back as an object rather than a list of one.
    decoded = cast("object", json.loads(done.stdout))
    rows = cast("list[Mapping[str, object]]", decoded if isinstance(decoded, list) else [decoded])
    out: dict[int, DaoRow] = {}
    for entry in rows:
        files = cast("list[Mapping[str, object]]", entry.get("files") or [])
        tags = cast("list[object]", entry.get("tags") or [])
        identifier = int(str(entry["id"]))
        out[identifier] = DaoRow(
            id=identifier,
            files=[
                (str(f["name"]), str(f["type"]), bytes.fromhex(str(f["hex"]))) for f in files
            ],
            tags=[str(t) for t in tags],
        )
    return out


@pytest.fixture
def database(tmp_path: Path) -> tuple[AccessDatabase, Path]:
    out = tmp_path / "complex.accdb"
    shutil.copyfile(FIXTURE, out)
    return AccessDatabase(out), out


def test_dao_reads_the_attachments_we_write(database: tuple[AccessDatabase, Path]) -> None:
    db, path = database
    things = db.table("Things")
    row = next(r for r in things.rows() if r["Id"] == 3)
    key = int(str(row["Files"]))
    payload = bytes(range(256)) * 4  # compressible, and not text
    things.set_attachments(
        "Files",
        key,
        [Attachment("written.dat", payload), Attachment("small.png", b"\x89PNG\r\n\x1a\n")],
    )
    db.save(path)

    # DAO walks the flat table's (key, FileName) index, so it reports
    # them by name; the reader orders them the same way.
    assert through_dao(path, "Things", "Files", "Tags")[3].files == [
        ("small.png", "png", b"\x89PNG\r\n\x1a\n"),
        ("written.dat", "dat", payload),
    ]
    assert [(a.name, a.type, a.data) for a in things.attachments("Files", key)] == [
        ("small.png", "png", b"\x89PNG\r\n\x1a\n"),
        ("written.dat", "dat", payload),
    ]


def test_dao_reads_the_multi_values_we_write(database: tuple[AccessDatabase, Path]) -> None:
    db, path = database
    things = db.table("Things")
    row = next(r for r in things.rows() if r["Id"] == 2)
    things.set_multi_values("Tags", int(str(row["Tags"])), ["one", "two", "three"])
    db.save(path)

    assert through_dao(path, "Things", "Files", "Tags")[2].tags == ["one", "three", "two"]


def test_dao_reads_a_row_we_inserted(database: tuple[AccessDatabase, Path]) -> None:
    """The complex id has to come from the counter at 0x1C: a row whose
    id collides with another row's is one the engine will not read."""
    db, path = database
    things = db.table("Things")
    things.insert_row({"Id": 7, "Name": "seventh"})
    row = next(r for r in things.rows() if r["Id"] == 7)
    things.set_attachments("Files", int(str(row["Files"])), [Attachment("new.txt", b"a new row")])
    things.set_multi_values("Tags", int(str(row["Tags"])), ["fresh"])
    db.save(path)

    reported = through_dao(path, "Things", "Files", "Tags")
    assert set(reported) == {1, 2, 3, 7}
    assert reported[7].files == [("new.txt", "txt", b"a new row")]
    assert reported[7].tags == ["fresh"]


def test_the_reader_agrees_with_dao_on_what_dao_wrote(
    database: tuple[AccessDatabase, Path]
) -> None:
    """The fixture's own rows, which DAO wrote, must decode to the bytes
    DAO hands back for them."""
    db, path = database
    things = db.table("Things")
    reported = through_dao(path, "Things", "Files", "Tags")

    for row in things.rows():
        identifier = int(str(row["Id"]))
        ours = things.attachments("Files", int(str(row["Files"])))
        assert [(a.name, a.type, a.data) for a in ours] == reported[identifier].files
        assert things.multi_values("Tags", int(str(row["Tags"]))) == reported[identifier].tags


def test_dao_reads_a_table_and_columns_we_created(tmp_path: Path) -> None:
    """The whole thing from nothing: our own database, our own table, an
    attachment column and a multi-valued one created by us, and the ACE
    engine reading the bytes back."""
    path = tmp_path / "created.accdb"
    db = AccessDatabase.create_new(path)
    notes = db.create_table(
        "Notes",
        [ColumnSpec("Id", "Long"), ColumnSpec("Title", "Text", size=60)],
        [IndexSpec("PrimaryKey", ("Id",), primary=True)],
    )
    notes.insert_row({"Id": 1, "Title": "first"})
    notes.insert_row({"Id": 2, "Title": "second"})
    notes.add_complex_column("Files", "attachment")
    notes.add_complex_column("Tags", "Text")
    payload = bytes(range(256))
    notes.set_attachments(
        "Files", 1, [Attachment("a.txt", b"first file"), Attachment("b.png", payload)]
    )
    notes.set_multi_values("Tags", 1, ["alpha", "beta"])
    notes.insert_row({"Id": 3, "Title": "third"})
    third = next(r for r in notes.rows() if r["Id"] == 3)
    notes.set_multi_values("Tags", int(str(third["Tags"])), ["gamma"])
    db.save(path)

    reported = through_dao(path, "Notes", "Files", "Tags")
    assert set(reported) == {1, 2, 3}
    assert reported[1].files == [
        ("a.txt", "txt", b"first file"),
        ("b.png", "png", payload),
    ]
    assert reported[1].tags == ["alpha", "beta"]
    assert reported[2].files == [] and reported[2].tags == []
    assert reported[3].tags == ["gamma"]


def test_compressed_attachments_match_the_engine_byte_for_byte(tmp_path: Path) -> None:
    """DAO attaches five files -- text, CSV, a small file, repetitive text
    and random bytes with runs -- and the FileData the engine stores for
    each, compressed with its own deflate, is exactly what
    ``encode_file_data`` produces: the same zlib header, blocks and
    trailer.  The deflate is classic zlib's at level 5, memLevel 7 and a
    32 KB window, carried in ``pyopenvba._deflate``."""
    import random

    random.seed(7)
    payloads = {
        "rep.txt": b"abcabcabcabc" * 400,
        "prose.txt": ("The quick brown fox jumps over the lazy dog. " * 80).encode()
        + "".join(random.choice("the engine writes rows on pages and pages on files ") for _ in range(4000)).encode(),
        "small.txt": b"hello world",
        "mixed.csv": b"".join(f"{i},{i * i},{random.random():.6f},name{i % 17}\r\n".encode() for i in range(3000)),
        "bytes.bin": bytes(random.randrange(256) for _ in range(20000)) + b"\x00" * 5000 + bytes(range(256)) * 40,
    }
    target = tmp_path / "attach.accdb"
    shutil.copy(HERE / "live_access_test" / "complex_columns.accdb", target)
    listing = tmp_path / "files.txt"
    lines: list[str] = []
    for n, (name, data) in enumerate(payloads.items(), start=100):
        (tmp_path / name).write_bytes(data)
        lines.append(f"Files{chr(9)}{n}{chr(9)}{tmp_path / name}")
    listing.write_text(chr(10).join(lines), encoding="utf-8")
    done = subprocess.run(
        ["powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(HERE / "live_access_test" / "dao_oracle.ps1"),
         "-Command", "attach-files", "-Path", str(target), "-Table", "Things", "-SqlFile", str(listing)],
        capture_output=True, text=True, timeout=_TIMEOUT,
    )
    assert done.returncode == 0 and done.stdout.strip() == "ok", f"DAO oracle failed: {done.stdout[-200:]} {done.stderr[-300:]}"
    db = AccessDatabase(target)
    column = next(c for c in db.complex_columns() if c.column == "Files")
    stored = {str(row["FileName"]): row["FileData"] for row in db.table(column.flat_table).rows()}
    problems: list[str] = []
    for name, data in payloads.items():
        engine = stored[name]
        assert isinstance(engine, bytes)
        ours = encode_file_data(name.rsplit(".", 1)[1], data)
        if ours != engine:
            first = next((i for i in range(min(len(ours), len(engine))) if ours[i] != engine[i]), min(len(ours), len(engine)))
            problems.append(f"{name}: ours {len(ours)} bytes, engine {len(engine)}, first difference at {first}")
    assert not problems, chr(10).join(problems)
