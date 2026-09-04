"""Live Access gate for macros (opt-in).

A macro blob that parses and rebuilds is still only a blob agreeing with
itself.  This gate hands what pyOpenVBA wrote to Access, **runs it**, and
reads back the value it set.

Opt-in: set ``RUN_LIVE_ACCESS_VBA=1`` on a Windows machine with desktop
Access and ``pyvbaharness`` installed.  ``pyvbaharness`` is a test-time
oracle only; pyOpenVBA never uses COM.
"""

from __future__ import annotations

import os
import shutil
import sys
from collections.abc import Callable
from pathlib import Path

import pytest

from pyopenvba.access import AccessDatabase, MacroAction

pytestmark = pytest.mark.skipif(
    os.environ.get("RUN_LIVE_ACCESS_VBA") != "1" or sys.platform != "win32",
    reason="live Access gate: set RUN_LIVE_ACCESS_VBA=1 on Windows with Access installed",
)

_TEMPLATE = (
    Path(__file__).parents[1]
    / "src"
    / "pyopenvba"
    / "_templates"
    / "blank_files"
    / "blank_database.accdb"
)
_TIMEOUT = 120.0

_PROBES = """
Public Function MacroNames() As Variant
    Dim i As Long, s As String
    For i = 0 To CurrentProject.AllMacros.Count - 1
        s = s & CurrentProject.AllMacros(i).Name & ";"
    Next i
    MacroNames = s
End Function

Public Function RunAndRead(ByVal name As String, ByVal variable As String) As Variant
    DoCmd.RunMacro name
    RunAndRead = TempVars(variable).Value
End Function

Public Function RunTwice(ByVal name As String, ByVal variable As String) As Variant
    DoCmd.RunMacro name
    DoCmd.RunMacro name
    RunTwice = TempVars(variable).Value
End Function
"""


def ask(path: Path, proc: str, *args: str) -> object:
    harness = pytest.importorskip("pyvbaharness")

    with harness.AccessSession() as access:
        access.open_document(path, read_only=False)
        result = access.run_vba(_PROBES, proc=proc, args=tuple(args), timeout=_TIMEOUT)
        assert result.outcome == "passed", (
            f"{proc}{tuple(args)}: {result.outcome} {getattr(result, 'error', None)}"
        )
        return result.value


@pytest.fixture
def blank(tmp_path: Path) -> Path:
    if not _TEMPLATE.exists():  # pragma: no cover - the template ships with the package
        pytest.skip("blank template not present")
    out = tmp_path / "blank.accdb"
    shutil.copyfile(_TEMPLATE, out)
    return out


def written(blank: Path, out: Path, build: Callable[[AccessDatabase], object]) -> Path:
    database = AccessDatabase(blank)
    build(database)
    database.save(out)
    return out


def test_access_lists_the_macros_we_write(blank: Path, tmp_path: Path) -> None:
    def build(db: AccessDatabase) -> None:
        db.create_macro("Ping", [MacroAction("Beep")])
        db.create_macro("Pong", [MacroAction("Beep"), MacroAction("Beep")])

    out = written(blank, tmp_path / "listed.accdb", build)

    assert {n for n in str(ask(out, "MacroNames")).split(";") if n} == {"Ping", "Pong"}


def test_access_runs_a_macro_we_write(blank: Path, tmp_path: Path) -> None:
    """`SetTempVar` leaves something to read, which is the only way to
    tell a macro that ran from one that merely loaded."""
    out = written(
        blank,
        tmp_path / "ran.accdb",
        lambda db: db.create_macro(
            "Probe", [MacroAction("SetTempVar", ("probe", "42")), MacroAction("Beep")]
        ),
    )

    assert ask(out, "RunAndRead", "Probe", "probe") == 42


def test_an_argument_that_is_an_expression_is_evaluated(blank: Path, tmp_path: Path) -> None:
    out = written(
        blank,
        tmp_path / "expression.accdb",
        lambda db: db.create_macro(
            "Sum", [MacroAction("SetTempVar", ("total", "6 * 7"))]
        ),
    )

    assert ask(out, "RunAndRead", "Sum", "total") == 42


def test_a_macro_of_several_actions_runs_them_in_order(blank: Path, tmp_path: Path) -> None:
    def build(db: AccessDatabase) -> None:
        db.create_macro(
            "Steps",
            [
                MacroAction("SetTempVar", ("step", "1")),
                MacroAction("SetTempVar", ("step", "2")),
                MacroAction("SetTempVar", ("step", "3")),
            ],
        )

    out = written(blank, tmp_path / "ordered.accdb", build)

    assert ask(out, "RunAndRead", "Steps", "step") == 3


def test_a_deleted_macro_is_gone_from_access(blank: Path, tmp_path: Path) -> None:
    def build(db: AccessDatabase) -> None:
        db.create_macro("Keep", [MacroAction("SetTempVar", ("kept", "7"))])
        db.create_macro("Drop", [MacroAction("Beep")])
        db.delete_macro("Drop")

    out = written(blank, tmp_path / "deleted.accdb", build)

    assert {n for n in str(ask(out, "MacroNames")).split(";") if n} == {"Keep"}
    assert ask(out, "RunAndRead", "Keep", "kept") == 7
