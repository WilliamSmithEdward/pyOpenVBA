"""Live ACE gate for the storage engine (opt-in).

The ACE engine itself, driven through DAO from PowerShell with no Access
window and no VBA, builds a table holding every column type and fills it.
The engine reads the file back with no Office involved.  The two views
must agree field for field, so what is compared is exactly what the real
engine would hand to any client.

Opt-in: set ``RUN_LIVE_ACCESS=1`` on a Windows machine with the Access
database engine installed (it ships with Office).  Skipped everywhere
else, including CI.  DAO is a test-time oracle only; pyOpenVBA never uses
COM.
"""

from __future__ import annotations

import datetime as dt
import json
import os
import shutil
import subprocess
import sys
import uuid
from decimal import Decimal
from pathlib import Path

import pytest

from pyopenvba.access import AccessDatabase

pytestmark = pytest.mark.skipif(
    os.environ.get("RUN_LIVE_ACCESS") != "1" or sys.platform != "win32",
    reason="live ACE gate: set RUN_LIVE_ACCESS=1 on Windows with Access installed",
)

HERE = Path(__file__).parent
ORACLE = HERE / "live_access_test" / "dao_oracle.ps1"
TEMPLATE = (
    HERE.parents[0]
    / "src"
    / "pyopenvba"
    / "_templates"
    / "blank_files"
    / "blank_database.accdb"
)
ROWS = 120
_TIMEOUT = 300


def oracle(*args: str) -> str:
    """Run the DAO oracle script; its stdout is the result."""
    done = subprocess.run(
        [
            "powershell",
            "-NoProfile",
            "-NonInteractive",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            str(ORACLE),
            *args,
        ],
        capture_output=True,
        timeout=_TIMEOUT,
        check=False,
    )
    stdout = done.stdout.decode("utf-8", errors="replace")
    if done.returncode != 0:
        raise AssertionError(
            f"dao_oracle.ps1 {' '.join(args)} failed ({done.returncode}):\n"
            f"{stdout}\n{done.stderr.decode('utf-8', errors='replace')}"
        )
    return stdout


def _format(value: object) -> str:
    """Mirror Format-Cell in dao_oracle.ps1 for the engine's decoded values."""
    if value is None:
        return "<null>"
    if isinstance(value, bool):
        return "True" if value else "False"
    if isinstance(value, Decimal):
        return f"{value:.4f}"
    if isinstance(value, float):
        text = f"{value:.6f}".rstrip("0").rstrip(".")
        return text if text not in ("", "-0") else "0"
    if isinstance(value, dt.datetime):
        return value.strftime("%Y-%m-%d %H:%M:%S")
    if isinstance(value, bytes):
        return value.hex()
    if isinstance(value, uuid.UUID):
        return "{" + str(value).upper() + "}"
    if isinstance(value, str):
        return value.replace("\t", "\\t").replace("\n", "\\n")
    return str(value)


def _engine_dump(db: AccessDatabase, name: str) -> list[str]:
    table = db.table(name)
    rows = sorted(table.rows(), key=lambda r: int(str(r["Id"])))
    return [
        "\t".join(f"{col}={_format(row[col])}" for col in table.column_names)
        for row in rows
    ]


def _diff(expected: str, mine: list[str]) -> list[str]:
    expected_lines = expected.split("\n")
    out = [
        f"row {i + 1}:\n  engine(ACE): {a}\n  pyopenvba:   {m}"
        for i, (a, m) in enumerate(zip(expected_lines, mine))
        if a != m
    ]
    if len(expected_lines) != len(mine):
        out.append(f"ACE returned {len(expected_lines)} rows, pyopenvba {len(mine)}")
    return out


def _check_indexes(db: AccessDatabase) -> None:
    """Every index the oracle created -- one per indexable column type,
    a descending one, a two-column one, a unique ignore-nulls one -- must
    decode to the row values it points at, in key order, with the node
    pages naming their children's last keys."""
    from pyopenvba.access._index import TextKey, node_pages, parse_index_page
    from pyopenvba.access._rows import split_row

    table = db.table("AllTypes")
    rows: dict[tuple[int, int], dict[str, object]] = {}
    for page, slot, data in table.raw_rows():
        rows[(page, slot)] = table.decode(split_row(table.definition, data))
    names = {i.name for i in table.indexes}
    assert {"IX_Txt", "IX_BigDesc", "IX_FlagTiny", "IX_UniqueBig", "IX_Frac", "IX_Uid", "IX_Bin"} <= names
    for index in table.indexes:
        expected = len(rows)
        if index.ignores_nulls:
            expected -= sum(
                1 for r in rows.values() if all(r[c.name] is None for c, _ in index.columns)
            )
        count = 0
        previous: list[object] | None = None
        for values, page, row in index.entries():
            count += 1
            actual = rows[(page, row)]
            for (column, _asc), value in zip(index.columns, values):
                if value is None:
                    assert actual[column.name] is None, (index.name, column.name)
                elif isinstance(value, TextKey):
                    assert isinstance(actual[column.name], str)
                elif isinstance(value, float):
                    assert isinstance(actual[column.name], float)
                    assert abs(value - float(str(actual[column.name]))) < 1e-9
                else:
                    assert value == actual[column.name], (index.name, column.name, value)
            previous = values
        assert previous is not None
        assert count == expected, (index.name, count, expected)
        for node in node_pages(db.store, index.real.root_page):
            for entry in node.entries:
                assert entry.child is not None
                last = parse_index_page(db.store, entry.child).entries[-1]
                assert (last.key, last.page, last.row) == (entry.key, entry.page, entry.row)
    assert table.index("IX_UniqueBig").unique and table.index("IX_UniqueBig").ignores_nulls
    assert table.index("IX_BigDesc").columns[0][1] is False


def test_every_column_type_reads_back_as_the_engine_shows_it(tmp_path: Path) -> None:
    target = tmp_path / "alltypes.accdb"
    shutil.copy(TEMPLATE, target)
    assert oracle("-Command", "build-alltypes", "-Path", str(target), "-Rows", str(ROWS)) == "ok"
    expected = json.loads(oracle("-Command", "dump", "-Path", str(target), "-Table", "AllTypes"))
    expected_wide = json.loads(oracle("-Command", "dump", "-Path", str(target), "-Table", "Wide"))
    assert isinstance(expected, str) and isinstance(expected_wide, str)

    db = AccessDatabase(target)
    table = db.table("AllTypes")
    assert table.row_count == ROWS
    assert table.column_names == [
        "Id", "Flag", "Tiny", "Small", "Big", "Cash", "Sgl", "Dbl", "Stamp",
        "Bin", "Txt", "Blob", "Story", "Uid", "Frac", "Huge",
    ]
    mismatches = _diff(expected, _engine_dump(db, "AllTypes"))
    assert not mismatches, f"{len(mismatches)} rows differ\n" + "\n".join(mismatches[:5])

    _check_indexes(db)

    # The wide table's definition spans pages.
    wide = db.table("Wide")
    assert len(wide.definition.pages) >= 2, wide.definition.pages
    assert wide.definition.definition_length > 4096
    assert len(wide.columns) == 151
    mismatches = _diff(expected_wide, _engine_dump(db, "Wide"))
    assert not mismatches, "\n".join(mismatches[:3])
