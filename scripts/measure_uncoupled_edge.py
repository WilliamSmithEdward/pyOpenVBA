"""What Excel and Word do with a TextBox whose BorderStyle and SpecialEffect are both set.

Microsoft's reference says a control takes its edge from one of the two,
and the designers clear one when the other is set. set_property writes the
field it names and nothing else, so on a TextBox, which is sunken by
default, set_property("BorderStyle", 1) leaves both set: a state the
designers never write. The library makes three forms, each with one TextBox
in the same place:

    Both      BorderStyle 1, SpecialEffect left sunken
    Border    BorderStyle 1 and SpecialEffect 0, the designers' result
    Effect    neither set: sunken, no border

Each application opens the file and, for each form at run time, reads the
TextBox's BorderStyle and SpecialEffect and has the form paint itself
(PrintWindow) to take the pixels of the TextBox's top-left corner. Then its
designer adds a Label to each form and saves, which rewrites the forms, and
the record keeps each TextBox's stored fields as the library reads them
back.

    python scripts/measure_uncoupled_edge.py

writes tests/fixtures/uncoupled_edge.json, which
tests/test_form_properties.py replays.
"""

from __future__ import annotations

import json
import tempfile
from pathlib import Path
from typing import Any

from pyvbaharness import ExcelSession, HarnessConfig, WordSession

from pyopenvba import ExcelFile, WordFile

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "tests" / "fixtures" / "uncoupled_edge.json"
FORMS = {"Both": {"BorderStyle": 1}, "Border": {"BorderStyle": 1, "SpecialEffect": 0}, "Effect": {}}
#: Where each form's TextBox sits, in points; the block taken, in pixels, from 8 left of and 4 below where
#: the client area ClientToScreen reports puts it, which in PrintWindow's picture of the window takes in the
#: TextBox's top-left corner.
LEFT, TOP, BLOCK, BEFORE_X, BEFORE_Y = 12, 12, 16, 8, -4

PROBE = "\r\n".join([
    "Option Explicit",
    "Private Type RECT",
    "    Left As Long",
    "    Top As Long",
    "    Right As Long",
    "    Bottom As Long",
    "End Type",
    "Private Type POINTAPI",
    "    X As Long",
    "    Y As Long",
    "End Type",
    'Private Declare PtrSafe Function FindWindowA Lib "user32" (ByVal c As String, ByVal t As String) As LongPtr',
    'Private Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal h As LongPtr, r As RECT) As Long',
    'Private Declare PtrSafe Function ClientToScreen Lib "user32" (ByVal h As LongPtr, p As POINTAPI) As Long',
    'Private Declare PtrSafe Function GetDC Lib "user32" (ByVal h As LongPtr) As LongPtr',
    'Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal h As LongPtr, ByVal dc As LongPtr) As Long',
    'Private Declare PtrSafe Function CreateCompatibleDC Lib "gdi32" (ByVal dc As LongPtr) As LongPtr',
    'Private Declare PtrSafe Function CreateCompatibleBitmap Lib "gdi32" (ByVal dc As LongPtr, '
    'ByVal w As Long, ByVal h As Long) As LongPtr',
    'Private Declare PtrSafe Function SelectObject Lib "gdi32" (ByVal dc As LongPtr, ByVal o As LongPtr) As LongPtr',
    'Private Declare PtrSafe Function DeleteObject Lib "gdi32" (ByVal o As LongPtr) As Long',
    'Private Declare PtrSafe Function DeleteDC Lib "gdi32" (ByVal dc As LongPtr) As Long',
    'Private Declare PtrSafe Function PrintWindow Lib "user32" (ByVal h As LongPtr, ByVal dc As LongPtr, '
    'ByVal flags As Long) As Long',
    'Private Declare PtrSafe Function GetPixel Lib "gdi32" (ByVal dc As LongPtr, ByVal x As Long, '
    'ByVal y As Long) As Long',
    "",
    "Private Function Corner(ByVal form As Object) As String",
    "    Dim h As LongPtr, r As RECT, p As POINTAPI, w As Long, ht As Long",
    "    Dim x As Long, y As Long, x0 As Long, y0 As Long",
    "    Dim sdc As LongPtr, mdc As LongPtr, bmp As LongPtr, old As LongPtr, out As String",
    "    form.Show vbModeless",
    "    form.Repaint",
    "    DoEvents",
    '    h = FindWindowA("ThunderDFrame", form.Caption)',
    "    GetWindowRect h, r",
    "    ClientToScreen h, p",
    "    w = r.Right - r.Left: ht = r.Bottom - r.Top",
    "    sdc = GetDC(0)",
    "    mdc = CreateCompatibleDC(sdc)",
    "    bmp = CreateCompatibleBitmap(sdc, w, ht)",
    "    old = SelectObject(mdc, bmp)",
    "    PrintWindow h, mdc, 0",
    f"    x0 = p.X - r.Left + {LEFT * 4 // 3 - BEFORE_X}: y0 = p.Y - r.Top + {TOP * 4 // 3 - BEFORE_Y}",
    f"    For y = 0 To {BLOCK - 1}",
    f"        For x = 0 To {BLOCK - 1}",
    '            out = out & Right$("00000" & Hex$(GetPixel(mdc, x0 + x, y0 + y)), 6)',
    "        Next",
    '        out = out & "/"',
    "    Next",
    "    SelectObject mdc, old",
    "    DeleteObject bmp",
    "    DeleteDC mdc",
    "    ReleaseDC 0, sdc",
    '    Corner = form.Box.BorderStyle & "|" & form.Box.SpecialEffect & "|" & out',
    "    Unload form",
    "End Function",
    "",
    "Public Function Measure() As String",
    '    Measure = Corner(Both) & ";" & Corner(Border) & ";" & Corner(Effect)',
    "End Function",
    "",
])

HOSTS: dict[str, dict[str, Any]] = {
    "excel": {"session": ExcelSession, "file": ExcelFile, "suffix": "xlsm", "open": "open_workbook",
              "opened": 'Workbooks.Open("{path}")', "quiet": "Application.DisplayAlerts = False"},
    "word": {"session": WordSession, "file": WordFile, "suffix": "docm", "open": "open_document",
             "opened": 'Documents.Open(FileName:="{path}")', "quiet": "Application.DisplayAlerts = 0"},
}


def make(path: Path, kind: Any) -> None:
    with kind.create_new(path) as document:
        # Every form is added before any is edited: add_form rereads the forms, which drops an earlier
        # form's unsaved edits.
        for name in FORMS:
            document.add_form(name, caption="Edge")
        for form in document.forms():
            box = form.add_control("TextBox", "Box", left=LEFT, top=TOP)
            for field, value in FORMS[form.name].items():
                box.set_property(field, value)
        document.vba_project().add_module("Probe", PROBE)
        document.save()


def stored(path: Path, kind: Any) -> dict[str, dict[str, object]]:
    with kind(path) as document:
        return {form.name: {field: form.control("Box").get(field) for field in ("BorderStyle", "SpecialEffect")}
                for form in document.forms()}


def run(office: Any, body: str) -> str:
    done = office.run_vba(f"Public Function Step() As String\nDim d As Object, out As String\n{body}\n"
                          "Step = out\nEnd Function\n", "Step", timeout=300.0)
    assert done.ok, f"{done.outcome} {done.message}"
    return str(done.value)


def measure(host: dict[str, Any], folder: Path) -> dict[str, Any]:
    path = folder / f"edge.{host['suffix']}"
    make(path, host["file"])
    written = stored(path, host["file"])
    with host["session"](HarnessConfig(lock_wait_s=1200.0)) as office:
        getattr(office, host["open"])(path, read_only=False)
        result = office.run_macro("Probe.Measure", timeout=180.0)
        assert result.ok, f"{result.outcome} {result.message}"
    shown = {}
    for name, part in zip(FORMS, str(result.value).split(";"), strict=True):
        border, effect, corner = part.split("|", 2)
        shown[name] = {"BorderStyle": int(border), "SpecialEffect": int(effect), "corner": corner.rstrip("/")}
    edit = [f'd.VBProject.VBComponents("{name}").Designer.Controls.Add "Forms.Label.1", "Touch"' for name in FORMS]
    with host["session"](HarnessConfig(lock_wait_s=1200.0)) as office:
        office.new_document()
        opened = host["opened"].format(path=path)
        if host["suffix"] == "xlsm":
            # Excel stops at a VBA dialog when one call changes a project and saves it: a call each.
            run(office, "\n".join([host["quiet"], f"Set d = {opened}", *edit, "out = d.Name"]))
            run(office, f'{host["quiet"]}\nSet d = Workbooks("{path.name}")\nd.Save\nd.Close False')
        else:
            run(office, "\n".join([host["quiet"], f"Set d = {opened}", *edit, "d.Save", "d.Close SaveChanges:=0"]))
    return {"written": written, "run_time": shown, "resaved": stored(path, host["file"])}


def main() -> None:
    folder = Path(tempfile.mkdtemp())
    record = {name: measure(host, folder) for name, host in HOSTS.items()}
    OUT.write_text(json.dumps(record, indent=1) + "\n", encoding="utf-8")
    print("wrote", OUT.relative_to(ROOT))


if __name__ == "__main__":
    main()
