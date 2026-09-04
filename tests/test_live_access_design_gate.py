"""Live Access gate for forms and reports (opt-in).

A design blob that parses and rebuilds is a blob agreeing with itself.
This gate hands what pyOpenVBA wrote to Access and **opens it in the
designer**, which is the only thing that says the design is one Access
will accept.

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

from pyopenvba.access import AccessDatabase

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
Public Function DesignNames() As Variant
    Dim i As Long, s As String
    For i = 0 To CurrentProject.AllForms.Count - 1
        s = s & "F:" & CurrentProject.AllForms(i).Name & ";"
    Next i
    For i = 0 To CurrentProject.AllReports.Count - 1
        s = s & "R:" & CurrentProject.AllReports(i).Name & ";"
    Next i
    DesignNames = s
End Function

Public Function OpenFormDesign(ByVal name As String) As Variant
    DoCmd.OpenForm name, acDesign
    OpenFormDesign = Forms(name).Name & "|" & Forms(name).Controls.Count & "|" & Forms(name).Section(0).Height
    DoCmd.Close acForm, name, acSaveNo
End Function

Public Function OpenReportDesign(ByVal name As String) As Variant
    DoCmd.OpenReport name, acViewDesign
    OpenReportDesign = Reports(name).Name & "|" & Reports(name).Controls.Count & "|" & Reports(name).Section(0).Height
    DoCmd.Close acReport, name, acSaveNo
End Function

Public Function DescribeControls(ByVal name As String) As Variant
    Dim c As Object, s As String
    DoCmd.OpenForm name, acDesign
    For Each c In Forms(name).Controls
        s = s & c.Name & ":" & TypeName(c) & ":" & c.Left & "," & c.Top & "," & c.Width & "," & c.Height
        If TypeName(c) = "Label" Then
            s = s & ":" & c.Caption
        ElseIf TypeName(c) = "TextBox" Then
            s = s & ":" & c.ControlSource
        End If
        s = s & ";"
    Next c
    DoCmd.Close acForm, name, acSaveNo
    DescribeControls = s
End Function

Public Function RunForm(ByVal name As String) As Variant
    DoCmd.OpenForm name
    RunForm = Forms(name).Name & "|" & Forms(name).CurrentView
    DoCmd.Close acForm, name, acSaveNo
End Function
"""


def ask(path: Path, proc: str, *args: str) -> object:
    harness = pytest.importorskip("pyvbaharness")

    with harness.AccessSession() as access:
        access.open_document(path, read_only=False)
        result = access.run_vba(_PROBES, proc=proc, args=tuple(args), timeout=_TIMEOUT)
        # A design Access will not open shows a dialog rather than
        # failing, so the dialog's own words are worth reporting.
        reported: list[object] = list(result.dialogs or [])
        dialogs = "; ".join(str(getattr(d, "message", "")) for d in reported)
        assert result.outcome == "passed", (
            f"{proc}{tuple(args)}: {result.outcome} {getattr(result, 'error', None)} {dialogs}"
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


def test_access_lists_the_designs_we_write(blank: Path, tmp_path: Path) -> None:
    def build(db: AccessDatabase) -> None:
        db.create_form("Plain")
        db.create_report("Sheet")

    out = written(blank, tmp_path / "listed.accdb", build)

    assert {n for n in str(ask(out, "DesignNames")).split(";") if n} == {"F:Plain", "R:Sheet"}


def test_access_opens_a_form_we_write_in_the_designer(blank: Path, tmp_path: Path) -> None:
    """A design Access will not open is one this gate exists to catch."""
    out = written(blank, tmp_path / "form.accdb", lambda db: db.create_form("Plain"))

    name, controls, height = str(ask(out, "OpenFormDesign", "Plain")).split("|")
    assert name == "Plain"
    assert int(controls) == 0
    assert int(height) > 0  # the Detail section has a size of its own


def test_access_opens_a_report_we_write_in_the_designer(blank: Path, tmp_path: Path) -> None:
    out = written(blank, tmp_path / "report.accdb", lambda db: db.create_report("Sheet"))

    name, controls, height = str(ask(out, "OpenReportDesign", "Sheet")).split("|")
    assert name == "Sheet"
    assert int(controls) == 0
    assert int(height) > 0


def test_a_form_we_write_runs(blank: Path, tmp_path: Path) -> None:
    """Opening it in form view, not just the designer."""
    out = written(blank, tmp_path / "run.accdb", lambda db: db.create_form("Plain"))

    name, _view = str(ask(out, "RunForm", "Plain")).split("|")
    assert name == "Plain"


def test_two_forms_both_open(blank: Path, tmp_path: Path) -> None:
    """The second takes a storage folder of its own, and Access will not
    find an object in a folder named anything else."""

    def build(db: AccessDatabase) -> None:
        db.create_form("First")
        db.create_form("Second")

    out = written(blank, tmp_path / "two.accdb", build)

    assert str(ask(out, "OpenFormDesign", "First")).startswith("First|")
    assert str(ask(out, "OpenFormDesign", "Second")).startswith("Second|")


def test_a_deleted_design_is_gone_from_access(blank: Path, tmp_path: Path) -> None:
    def build(db: AccessDatabase) -> None:
        db.create_form("Keep")
        db.create_form("Drop")
        db.create_report("Sheet")
        db.delete_form("Drop")
        db.delete_report("Sheet")

    out = written(blank, tmp_path / "deleted.accdb", build)

    assert {n for n in str(ask(out, "DesignNames")).split(";") if n} == {"F:Keep"}
    assert str(ask(out, "OpenFormDesign", "Keep")).startswith("Keep|")


def test_access_reads_back_the_controls_we_add(blank: Path, tmp_path: Path) -> None:
    """Name, type and every measurement, as Access reports them."""

    def build(db: AccessDatabase) -> None:
        db.create_form("Built")
        db.add_control(
            "Built", "Label", "Title", left=240, top=240, width=2000, height=300,
            caption="Hello there",
        )
        db.add_control(
            "Built", "TextBox", "Box", left=240, top=700, width=2400, height=320,
            caption="=1+1",
        )

    out = written(blank, tmp_path / "controls.accdb", build)

    reported = [c for c in str(ask(out, "DescribeControls", "Built")).split(";") if c]
    assert reported == [
        "Title:Label:240,240,2000,300:Hello there",
        "Box:TextBox:240,700,2400,320:=1+1",
    ]


def test_controls_on_a_report_reach_access(blank: Path, tmp_path: Path) -> None:
    """Three across two sections, which is what makes the markers matter:
    the page header holds one and the detail band two, and Access refuses
    either encoding in the other's place."""

    def build(db: AccessDatabase) -> None:
        db.create_report("Sheet")
        db.add_control("Sheet", "Label", "Heading", kind="report", caption="Monthly", width=2000)
        db.add_control("Sheet", "TextBox", "Total", kind="report", caption="=2*21", top=500)
        db.add_control(
            "Sheet", "Label", "PageTitle", kind="report", section="PageHeaderSection",
            caption="Header band", width=2000,
        )

    out = written(blank, tmp_path / "report_control.accdb", build)

    name, controls, _height = str(ask(out, "OpenReportDesign", "Sheet")).split("|")
    assert name == "Sheet" and int(controls) == 3


def test_one_control_on_a_form_reaches_access(blank: Path, tmp_path: Path) -> None:
    """A lone control takes the other marker, so it needs its own check."""
    out = written(
        blank,
        tmp_path / "one.accdb",
        lambda db: (db.create_form("Solo"), db.add_control("Solo", "Label", "Only", caption="Hi")),
    )

    assert [c for c in str(ask(out, "DescribeControls", "Solo")).split(";") if c] == [
        "Only:Label:0,0,1440,240:Hi"
    ]
