"""Live Access gate for the VBA module writers (opt-in).

Everything else about a written module can be checked from the file and
still be wrong: the dir stream can list a module Access refuses to name,
the storage rows can be complete and the folder still be called something
Access will not look under.  Three defects were caught only here, and all
three had agreed with Access by accident on every fixture used before --
the storage folder's name, the `MSysObjects` id step, and a folder left
behind by a delete.  So this gate asks Access itself, and asks for the
strongest answer available: **run the code and compare the value it
returns**.

Opt-in: set ``RUN_LIVE_ACCESS_VBA=1`` on a Windows machine with desktop
Access and ``pyvbaharness`` installed.  Skipped everywhere else,
including CI.  ``pyvbaharness`` is a test-time oracle only; pyOpenVBA
never uses COM.
"""

from __future__ import annotations

import os
import shutil
import sys
from collections.abc import Callable
from pathlib import Path

import pytest

from pyopenvba.access import AccessDatabase

pytestmark = pytest.mark.skipif(
    os.environ.get("RUN_LIVE_ACCESS_VBA") != "1" or sys.platform != "win32",
    reason="live Access VBA gate: set RUN_LIVE_ACCESS_VBA=1 on Windows with Access installed",
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

#: Injected into the database under test so every probe runs inside
#: Access's own VBA.  Over a bare COM boundary Access refuses half of
#: these verbs, and a failure there would say nothing.
_PROBES = """
Public Function CallProc(ByVal name As String) As Variant
    CallProc = Application.Run(name)
End Function

Public Function AccessModules() As Variant
    Dim i As Long, s As String
    For i = 0 To CurrentProject.AllModules.Count - 1
        s = s & CurrentProject.AllModules(i).Name & ";"
    Next i
    AccessModules = s
End Function

Public Function ReadLines(ByVal moduleName As String) As Variant
    Dim c As Object
    Set c = Application.VBE.ActiveVBProject.VBComponents(moduleName)
    ReadLines = c.CodeModule.Lines(1, c.CodeModule.CountOfLines)
End Function

Public Function AddProcedure(ByVal moduleName As String) As Variant
    Dim c As Object
    Set c = Application.VBE.ActiveVBProject.VBComponents(moduleName)
    c.CodeModule.AddFromString "Public Function Probe9() As Variant" & vbCrLf & _
        "    Probe9 = 999" & vbCrLf & "End Function"
    AddProcedure = c.CodeModule.CountOfLines
End Function
"""


def ask(path: Path, proc: str, *args: str) -> object:
    """Open the database in Access and run one probe against it."""
    harness = pytest.importorskip("pyvbaharness")

    with harness.AccessSession() as access:
        access.open_document(path, read_only=False)
        result = access.run_vba(_PROBES, proc=proc, args=tuple(args), timeout=_TIMEOUT)
        assert result.outcome == "passed", (
            f"{proc}{tuple(args)}: {result.outcome} {getattr(result, 'error', None)}"
        )
        return result.value


def modules(path: Path) -> set[str]:
    """The names `CurrentProject.AllModules` reports.  Its order is not
    the order modules were added, so only the set is asserted."""
    return {name for name in str(ask(path, "AccessModules")).split(";") if name}


@pytest.fixture
def blank(tmp_path: Path) -> Path:
    """A fresh copy of the shipped template, which holds one module."""
    if not _TEMPLATE.exists():  # pragma: no cover - the template ships with the package
        pytest.skip("blank template not present")
    out = tmp_path / "blank.accdb"
    shutil.copyfile(_TEMPLATE, out)
    return out


ADDER = (
    "Option Compare Database\n"
    "\n"
    "Public Function AdderGo() As Variant\n"
    "    AdderGo = 4242\n"
    "End Function"
)


def written(blank: Path, out: Path, build: Callable[[AccessDatabase], object]) -> Path:
    """Apply `build` to the template and save it under `out`."""
    database = AccessDatabase(blank)
    build(database)
    database.save(out)
    return out


class TestAccessRunsWhatWeCreate:
    def test_a_created_module_is_listed_and_its_code_runs(
        self, blank: Path, tmp_path: Path
    ) -> None:
        out = written(blank, tmp_path / "created.accdb", lambda db: db.create_module("Adder", ADDER))

        assert modules(out) == {"Module1", "Adder"}
        assert ask(out, "CallProc", "AdderGo") == 4242
        # the VBE reports the source without the attribute lines
        read_back = str(ask(out, "ReadLines", "Adder")).replace(chr(13) + chr(10), chr(10))
        assert read_back.strip() == ADDER.strip()

    def test_two_creates_without_access_in_between(self, blank: Path, tmp_path: Path) -> None:
        def build(db: AccessDatabase) -> None:
            db.create_module("Adder", ADDER)
            db.create_module(
                "Doubler",
                "Public Function DoublerGo() As Variant\n    DoublerGo = 77\nEnd Function",
            )

        out = written(blank, tmp_path / "two.accdb", build)

        assert modules(out) == {"Module1", "Adder", "Doubler"}
        assert ask(out, "CallProc", "DoublerGo") == 77

    def test_a_created_class_can_be_instantiated(self, blank: Path, tmp_path: Path) -> None:
        def build(db: AccessDatabase) -> None:
            db.create_module(
                "Widget",
                "Option Compare Database\n"
                "\n"
                "Private spun As Long\n"
                "\n"
                "Public Function Spin() As Variant\n"
                "    spun = spun + 5\n"
                "    Spin = spun\n"
                "End Function",
                kind="class",
            )
            db.create_module(
                "Driver",
                "Option Compare Database\n"
                "\n"
                "Public Function UseWidget() As Variant\n"
                "    Dim w As Widget\n"
                "    Set w = New Widget\n"
                "    UseWidget = w.Spin() + w.Spin()\n"
                "End Function",
            )

        out = written(blank, tmp_path / "widget.accdb", build)

        assert ask(out, "CallProc", "UseWidget") == 15

    def test_the_project_still_takes_an_edit_afterwards(
        self, blank: Path, tmp_path: Path
    ) -> None:
        """A created project that Access will not let you edit is the
        failure this whole route exists to avoid."""
        out = written(blank, tmp_path / "edited.accdb", lambda db: db.create_module("Adder", ADDER))

        assert int(str(ask(out, "AddProcedure", "Module1"))) > 0


class TestRenameAndDelete:
    def test_a_created_module_can_be_renamed(self, blank: Path, tmp_path: Path) -> None:
        def build(db: AccessDatabase) -> None:
            db.create_module("Adder", ADDER)
            db.rename_module("Adder", "Summer")

        out = written(blank, tmp_path / "renamed.accdb", build)

        assert modules(out) == {"Module1", "Summer"}
        assert ask(out, "CallProc", "AdderGo") == 4242

    def test_a_created_module_can_be_deleted(self, blank: Path, tmp_path: Path) -> None:
        def build(db: AccessDatabase) -> None:
            db.create_module("Adder", ADDER)
            db.delete_module("Adder")

        out = written(blank, tmp_path / "deleted.accdb", build)

        assert modules(out) == {"Module1"}

    def test_a_module_created_after_a_delete_still_resolves(
        self, blank: Path, tmp_path: Path
    ) -> None:
        """A delete that leaves the storage folder behind makes the next
        create pick a name Access will not look under, and `AllModules`
        then fails on it while the VBE still runs the code."""

        def build(db: AccessDatabase) -> None:
            db.create_module("First", ADDER)
            db.create_module("Second", "Option Compare Database")
            db.delete_module("First")
            db.create_module(
                "Third",
                "Public Function ThirdGo() As Variant\n    ThirdGo = 33\nEnd Function",
            )

        out = written(blank, tmp_path / "recycled.accdb", build)

        assert modules(out) == {"Module1", "Second", "Third"}
        assert ask(out, "CallProc", "ThirdGo") == 33


class TestReplacingSource:
    def test_source_a_pcode_writer_could_not_emit_still_runs(
        self, blank: Path, tmp_path: Path
    ) -> None:
        """`Const`, a module-level array, `Static` and a fixed-length
        string, none of which a p-code writer can produce."""
        code = (
            "Option Compare Database\n"
            "Option Explicit\n"
            "\n"
            'Private Const Greeting As String = "hi"\n'
            "Private Cache(1 To 3) As Long\n"
            "\n"
            "Public Function Grown() As Variant\n"
            "    Static calls As Long\n"
            "    Dim total As Long, i As Long\n"
            "    Dim label As String * 5\n"
            "    label = Greeting\n"
            "    calls = calls + 1\n"
            "    For i = 1 To 3\n"
            "        Cache(i) = i * i\n"
            "        total = total + Cache(i)\n"
            "    Next i\n"
            '    Grown = total & "/" & calls & "/" & Trim(label)\n'
            "End Function"
        )
        out = written(
            blank, tmp_path / "replaced.accdb", lambda db: db.set_module_source("Module1", code)
        )

        assert ask(out, "CallProc", "Grown") == "14/1/hi"
