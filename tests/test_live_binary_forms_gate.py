"""Live gate: UserForms in a binary .xls and .doc, both ways (opt-in).

What only Excel and Word can answer: that a form add_form writes into
their binary file opens, shows its controls and compiles, and that a form
their own editor inserts is one the library reads and edits, the edit
showing when the application opens the file again. Each file starts as
tests/fixtures/binary_project/, which the application saved itself. The
applications' compile check runs through pyvbaharness, as xlide_mcp's
does.

Opt-in per application: RUN_LIVE_EXCEL=1 or RUN_LIVE_WORD=1, on Windows
with it installed. Skipped everywhere else, including CI.
"""

from __future__ import annotations

import os
import shutil
import sys
from pathlib import Path
from typing import Any

import pytest

from pyopenvba import ExcelFile, VBAProjectError, WordFile

FIXTURES = Path(__file__).parent / "fixtures" / "binary_project"
#: Each application: the switch that opts in, the library's file type, the fixture, the harness session, the
#: open file as its VBA names it, and how its VBA opens, saves and closes a file.
HOSTS: dict[str, tuple[str, Any, str, str, str, str, str]] = {
    "excel": ("RUN_LIVE_EXCEL", ExcelFile, "workbook.xls", "ExcelSession", "ActiveWorkbook",
              'Application.DisplayAlerts = False\nSet d = Workbooks.Open("{path}")',
              'Application.DisplayAlerts = False\nSet d = Workbooks("{name}")\nd.Save\nd.Close False'),
    "word": ("RUN_LIVE_WORD", WordFile, "document.doc", "WordSession", "ActiveDocument",
             'Application.DisplayAlerts = 0\nSet d = Documents.Open(FileName:="{path}", AddToRecentFiles:=False)',
             "d.Save\nd.Close SaveChanges:=0"),
}
HANDLER = "Private Sub UserForm_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)\r\n    KeyAscii = 0\r\nEnd Sub\r\n"
MODULE = ("Public Sub ShowIt()\r\n    Dim f As MSForms.UserForm\r\n    Set f = Wizard\r\n"
          "    Dim c As MSForms.Control\r\n    For Each c In Wizard.Controls\r\n    Next\r\nEnd Sub\r\n")


def _param(host: str) -> Any:
    switch = HOSTS[host][0]
    return pytest.param(host, id=host, marks=pytest.mark.skipif(
        os.environ.get(switch) != "1" or sys.platform != "win32",
        reason=f"live binary forms gate: set {switch}=1 on Windows with the application installed"))


def _session(host: str) -> Any:
    harness = pytest.importorskip("pyvbaharness")
    return getattr(harness, HOSTS[host][3])(harness.HarnessConfig(lock_wait_s=1200.0))


def _run(office: Any, body: str) -> None:
    done = office.run_vba(f"Public Function Step() As String\nDim d As Object, c As Object\n{body}\n"
                          'Step = "ok"\nEnd Function\n', "Step", timeout=120.0)
    assert done.ok, f"{done.outcome}: {done.message}"


def _seen(host: str, target: Path, name: str, expressions: list[str]) -> list[object]:
    """What the application reports for each expression on component ``name``, the file opened afresh."""
    component = f'{HOSTS[host][4]}.VBProject.VBComponents("{name}")'
    with _session(host) as office:
        office.open_document(target, read_only=False)
        return [office.eval(expression.replace("VBC", component, 1)) for expression in expressions]


@pytest.mark.parametrize("host", [_param(host) for host in HOSTS])
def test_a_form_the_library_adds_opens_and_compiles(host: str, tmp_path: Path) -> None:
    _, opener, name, *_ = HOSTS[host]
    target = tmp_path / name
    shutil.copyfile(FIXTURES / name, target)
    with opener(target) as office_file:
        form = office_file.add_form("Wizard", caption="Setup")
        form.add_control("CommandButton", "Go", left=12, top=12)
        form.add_control("Frame", "Group", left=12, top=48, width=200, height=90)
        form.add_control("OptionButton", "First", container="Group", left=6, top=12)
        office_file.set_module("Wizard", HANDLER)
        office_file.vba_project().add_module("Shows", MODULE)
        office_file.save()
    assert _seen(host, target, "Wizard", [
        "VBC.Type", "VBC.Designer.Caption", "VBC.Designer.Controls.Count",
        'VBC.Designer.Controls("First").Parent.Name']) == [3, "Setup", 3, "Group"]
    with _session(host) as office:
        office.open_document(target)
        compiled = office.compile_project()
    assert compiled.ok, f"{compiled.outcome}: {compiled.message} {compiled.dialog}"


@pytest.mark.parametrize("host", [_param(host) for host in HOSTS])
def test_a_form_the_editor_adds_is_read_and_edited(host: str, tmp_path: Path) -> None:
    _, opener, name, _, _, opening, closing = HOSTS[host]
    target = tmp_path / name
    shutil.copyfile(FIXTURES / name, target)
    adding = ('Set c = d.VBProject.VBComponents.Add(3)\nc.Designer.Caption = "Editor"\n'
              'c.Designer.Controls.Add "Forms.CommandButton.1", "Go"')
    with _session(host) as office:
        office.new_document()
        if host == "excel":
            # Excel stops at a VBA dialog when one call changes a project and saves it: a call each.
            _run(office, f"{opening.format(path=target)}\n{adding}")
            _run(office, closing.format(name=name))
        else:
            # Word loses the harness's procedures once another file's project changed: one call.
            _run(office, f"{opening.format(path=target)}\n{adding}\n{closing}")
    with opener(target) as office_file:
        forms = office_file.forms()
        assert [(form.name, form.get("Caption"), [c.name for c in form.walk()]) for form in forms] == [
            ("UserForm1", "Editor", ["Go"])]
        with pytest.raises(VBAProjectError, match="required by UserForms: UserForm1"):
            office_file.remove_reference("MSForms")
        forms[0].set_property("Caption", "Edited")
        forms[0].add_control("Label", "Note", left=12, top=48, width=120)
        office_file.save()
    assert _seen(host, target, "UserForm1", [
        "VBC.Designer.Caption", "VBC.Designer.Controls.Count",
        'TypeName(VBC.Designer.Controls("Note"))']) == ["Edited", 2, "Label"]
