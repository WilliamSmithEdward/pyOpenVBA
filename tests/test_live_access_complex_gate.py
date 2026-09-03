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

from pyopenvba.access import AccessDatabase, Attachment

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
