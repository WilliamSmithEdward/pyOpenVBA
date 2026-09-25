"""Live gate: controls the library adds are what Excel's and Word's designers report (opt-in).

Each application makes a UserForm set to Arial 10 with a Frame set to
Courier New 9, and saves it. The library adds controls of each kind the
form designer fixes of #31 touch; the application opens the file and its
designer reports each new control's font, alignment, effect, pages, tabs,
tab order and height as tests/fixtures/form_designer.json and
form_fonts.json say the designer gives them, and the project compiles.

Opt-in per application: RUN_LIVE_EXCEL=1 or RUN_LIVE_WORD=1, on Windows
with it installed. Skipped everywhere else, including CI.
"""

from __future__ import annotations

import os
import sys
from pathlib import Path
from typing import Any

import pytest

from pyopenvba import ExcelFile, WordFile

#: Each application: the switch that opts in, the library's file type, the extension, the harness session,
#: the open file as its VBA names it, and how its VBA makes, saves and closes a file.
HOSTS: dict[str, tuple[str, Any, str, str, str, str, str]] = {
    "excel": ("RUN_LIVE_EXCEL", ExcelFile, "xlsm", "ExcelSession", "ActiveWorkbook",
              "Application.DisplayAlerts = False\nSet d = Workbooks.Add(-4167)",
              'Application.DisplayAlerts = False\nSet d = Workbooks("{name}")\n'
              'd.SaveAs Filename:="{path}", FileFormat:=52\nd.Close False'),
    "word": ("RUN_LIVE_WORD", WordFile, "docm", "WordSession", "ActiveDocument",
             "Application.DisplayAlerts = 0\nSet d = Documents.Add",
             'd.SaveAs2 FileName:="{path}", FileFormat:=13\nd.Close SaveChanges:=0'),
}
MAKING = ('Set c = d.VBProject.VBComponents.Add(3)\nc.Name = "Fonted"\nSet f = c.Designer\n'
          'f.Font.Name = "Arial"\nf.Font.Size = 10\nSet g = f.Controls.Add("Forms.Frame.1", "Frame1")\n'
          'g.Font.Name = "Courier New"\ng.Font.Size = 9\nout = d.Name')
SHOWS = ("Public Function Captions() As String\r\n    Fonted.Show vbModeless\r\n"
         "    Captions = Fonted.Caption\r\n    Unload Fonted\r\nEnd Function\r\n")
#: What the application reports for each expression on the form's designer, "C" standing for its Controls.
EXPECTED: list[tuple[str, object]] = [
    ('C("AddedLabel").Font.Name', "Arial"),
    ('C("AddedLabel").Font.Size', 9.75),
    ('C("InFrame").Font.Name', "Courier New"),
    ('C("InFrame").Font.Size', 9),
    ('C("AddedToggle").TextAlign', 2),       # fmTextAlignCenter
    ('C("AddedFrame").SpecialEffect', 3),    # fmSpecialEffectEtched
    ('C("AddedPages").Pages.Count', 2),
    ('C("AddedPages").Pages(1).Caption', "Page2"),
    ('C("AddedTabs").Tabs.Count', 2),
    ('C("AddedTabs").Tabs(1).Caption', "Tab2"),
    ('C("AddedCheck").Height', 19.5),       # 26 pixels: four more than a tab in Arial 9.75 pt
]


def _param(host: str) -> Any:
    switch = HOSTS[host][0]
    return pytest.param(host, id=host, marks=pytest.mark.skipif(
        os.environ.get(switch) != "1" or sys.platform != "win32",
        reason=f"live form designer gate: set {switch}=1 on Windows with the application installed"))


def _run(office: Any, body: str) -> str:
    done = office.run_vba("Public Function Step() As String\nDim d As Object, c As Object, f As Object, "
                          f"g As Object, out As String\n{body}\nStep = out\nEnd Function\n", "Step", timeout=300.0)
    assert done.ok, f"{done.outcome}: {done.message}"
    return str(done.value)


@pytest.mark.parametrize("host", [_param(host) for host in HOSTS])
def test_the_designer_reports_what_the_library_added(host: str, tmp_path: Path) -> None:
    harness = pytest.importorskip("pyvbaharness")
    _, opener, suffix, session_name, active, making, saving = HOSTS[host]
    target = tmp_path / f"designer.{suffix}"
    with getattr(harness, session_name)(harness.HarnessConfig(lock_wait_s=1200.0)) as office:
        office.new_document()
        if host == "excel":
            # Excel stops at a VBA dialog when one call changes a project and saves it: a call each.
            name = _run(office, f"{making}\n{MAKING}")
            _run(office, saving.format(name=name, path=target))
        else:
            # Word loses the harness's procedures once another file's project changed: one call.
            _run(office, f"{making}\n{MAKING}\n{saving.format(path=target)}")

    with opener(target) as office_file:
        form = next(form for form in office_file.forms() if form.name == "Fonted")
        for kind, name, container in [("Label", "AddedLabel", None), ("ToggleButton", "AddedToggle", None),
                                      ("Label", "InFrame", "Frame1"), ("Frame", "AddedFrame", None),
                                      ("MultiPage", "AddedPages", None), ("TabStrip", "AddedTabs", None),
                                      ("Image", "AddedImage", None), ("CommandButton", "AddedButton", None),
                                      ("CheckBox", "AddedCheck", None)]:
            form.add_control(kind, name, container=container, left=0, top=0)
        # The Image stays last in the tab order: the button added after it takes its place.
        stored = {control.name: control.tab_index for control in form.controls}
        assert stored["AddedImage"] == max(index or 0 for index in stored.values())
        # The caption a running form shows is its header's, which setting Caption writes with the record's.
        form.set_property("Caption", "Shown caption")
        office_file.vba_project().add_module("Shows", SHOWS)
        office_file.save()

    controls = f'{active}.VBProject.VBComponents("Fonted").Designer.Controls'
    with getattr(harness, session_name)(harness.HarnessConfig(lock_wait_s=1200.0)) as office:
        office.open_document(target, read_only=False)
        seen = [office.eval(expression.replace("C(", f"{controls}(", 1)) for expression, _ in EXPECTED]
        tab_indexes = {name: office.eval(f'{controls}("{name}").TabIndex')
                       for name in ("AddedButton", "AddedCheck", "AddedToggle")}
        designer_caption = office.eval(f'{active}.VBProject.VBComponents("Fonted").Designer.Caption')
        shown = office.run_macro("Shows.Captions")
        compiled = office.compile_project()
    assert [value for _, value in EXPECTED] == seen
    assert tab_indexes == {name: stored[name] for name in tab_indexes}
    assert (designer_caption, shown.value) == ("Shown caption", "Shown caption")
    assert compiled.ok, f"{compiled.outcome}: {compiled.message} {compiled.dialog}"
