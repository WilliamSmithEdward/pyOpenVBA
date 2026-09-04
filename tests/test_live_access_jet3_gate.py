"""Live Jet 3 gate: the Access 97 reader against the engine (opt-in).

Access dropped Jet 3 in 2013, but Jet 4.0 ships with Windows and still
creates and reads it through DAO 3.6.  Both are 32-bit, so the oracle runs
under SysWOW64 PowerShell.  It builds a database holding every column type
Jet 3 has, a memo too long for a 2 KiB page, text with code page
characters, a deleted row and four hundred rows spread over many pages;
pyOpenVBA reads the same file with no COM involved, and the two views must
agree cell for cell.

Opt-in: set ``RUN_LIVE_ACCESS_JET3=1`` on Windows with DAO 3.6 present
(``Common Files/Microsoft Shared/DAO/dao360.dll``).  DAO is a test-time
oracle only; pyOpenVBA never uses COM.
"""

from __future__ import annotations

import datetime as dt
import json
import os
import subprocess
import sys
import uuid
from decimal import Decimal
from pathlib import Path

import pytest

from pyopenvba.access import AccessDatabase

HERE = Path(__file__).parent
ORACLE = HERE / "live_access_test" / "dao_jet3.ps1"
#: DAO 3.6 is 32-bit, so the 64-bit shell cannot create it.
POWERSHELL32 = Path(os.environ.get("SystemRoot", r"C:\Windows")) / "SysWOW64" / "WindowsPowerShell" / "v1.0" / "powershell.exe"
_TIMEOUT = 300

pytestmark = pytest.mark.skipif(
    os.environ.get("RUN_LIVE_ACCESS_JET3") != "1" or sys.platform != "win32",
    reason="live Jet 3 gate: set RUN_LIVE_ACCESS_JET3=1 on Windows with DAO 3.6",
)


def oracle(*args: str) -> str:
    """Run the Jet 3 oracle under the 32-bit shell; its stdout is the result."""
    done = subprocess.run(
        [str(POWERSHELL32), "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(ORACLE), *args],
        capture_output=True,
        # The script sets its console to UTF-8; without saying so here
        # the code page characters the fixture carries arrive mangled.
        encoding="utf-8",
        timeout=_TIMEOUT,
    )
    if done.returncode != 0:
        raise AssertionError(
            f"dao_jet3.ps1 {' '.join(args)} failed ({done.returncode}):\n"
            f"{done.stdout}\n{done.stderr}"
        )
    return done.stdout.strip()


def _format(value: object) -> str:
    """Mirror Format-Cell in dao_jet3.ps1."""
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
        return value.replace("\t", "\t").replace("\n", "\n")
    return str(value)


def _agree(theirs: str, mine: str) -> bool:
    """Cells match outright, or -- for a float near the type's limit, where
    the two runtimes print different numbers of significant digits -- to
    within a relative whisker."""
    if theirs == mine:
        return True
    try:
        a, b = float(theirs), float(mine)
    except ValueError:
        return False
    if a == b:
        return True
    return abs(a - b) <= 1e-6 * max(abs(a), abs(b))


def _rows(dumped: str) -> list[list[str]]:
    return [line.split("\t") for line in dumped.split("\n") if line]


def _ours(db: AccessDatabase, name: str) -> list[list[str]]:
    table = db.table(name)
    rows = sorted(table.rows(), key=lambda r: int(str(r["Id"])))
    return [[f"{c}={_format(row[c])}" for c in table.column_names] for row in rows]


def _differences(theirs: list[list[str]], mine: list[list[str]]) -> list[str]:
    out: list[str] = []
    if len(theirs) != len(mine):
        out.append(f"row count: engine {len(theirs)}, pyopenvba {len(mine)}")
    for i, (their_row, my_row) in enumerate(zip(theirs, mine)):
        if len(their_row) != len(my_row):
            out.append(f"row {i}: engine has {len(their_row)} cells, pyopenvba {len(my_row)}")
            continue
        for their_cell, my_cell in zip(their_row, my_row):
            their_name, _, their_value = their_cell.partition("=")
            my_name, _, my_value = my_cell.partition("=")
            if their_name != my_name:
                out.append(f"row {i}: engine names {their_name}, pyopenvba {my_name}")
            elif not _agree(their_value, my_value):
                out.append(f"row {i} {their_name}: engine {their_value!r}, pyopenvba {my_value!r}")
    return out


@pytest.fixture(scope="module")
def built(tmp_path_factory: pytest.TempPathFactory) -> Path:
    path = tmp_path_factory.mktemp("jet3") / "gate.mdb"
    assert oracle("-Command", "build", "-Path", str(path)) == "ok"
    return path


@pytest.mark.parametrize("name", ["AllTypes", "Many"])
def test_the_reader_sees_what_the_engine_sees(built: Path, name: str) -> None:
    dumped = json.loads(oracle("-Command", "dump", "-Path", str(built), "-Table", name))
    assert isinstance(dumped, str)
    db = AccessDatabase(built)
    assert _differences(_rows(dumped), _ours(db, name)) == []


def test_the_hard_values_are_really_there(built: Path) -> None:
    """A dump that agreed on nothing interesting would still pass above, so
    name what the fixture was built to exercise."""
    db = AccessDatabase(built)
    rows = sorted(db.table("AllTypes").rows(), key=lambda r: int(str(r["Id"])))
    assert len(rows) == 4                      # one of the five was deleted
    assert rows[2]["T"] == "accented: caf\u00e9 na\u00efve"
    memo = rows[2]["M"]
    assert isinstance(memo, str) and len(memo) == 4000
    assert rows[3]["T"] is None and rows[3]["Flag"] is False
    assert len(db.table("Many").data_pages()) > 1
    assert db.store.page_size == 2048
