"""Have Office build the shapes, then keep what it wrote.

Nothing here guesses at DrawingML.  Each host is asked, through its own
object model, to add every kind of shape pyOpenVBA means to handle, and
to say what it then reports about each one: name, type, position, size,
text, and the macro a click runs.  The file Office saved becomes a
fixture and the answers become ``measured.json`` beside it, so the
reader and the writer are both held to what Office actually does.

Run it on a Windows machine with Office installed:

    python scripts/measure_shapes.py excel
    python scripts/measure_shapes.py powerpoint
    python scripts/measure_shapes.py word
    python scripts/measure_shapes.py all
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
FOLDER = ROOT / "tests" / "fixtures" / "shapes"

#: Excel: one of every kind a sheet can hold, each placed where its
#: anchor can be read back, and two of them wired to a macro.
EXCEL_BUILD = """
Sub Build()
    Dim s As Object, sh As Object
    Set s = ActiveWorkbook.Worksheets(1)
    Set sh = s.Shapes.AddShape(1, 10, 20, 100, 50)      ' msoShapeRectangle
    sh.Name = "Rect"
    sh.TextFrame2.TextRange.Text = "Hello"
    sh.OnAction = "Clicked"
    Set sh = s.Shapes.AddShape(5, 150, 20, 80, 80)      ' msoShapeRoundedRectangle
    sh.Name = "Rounded"
    Set sh = s.Shapes.AddShape(9, 250, 20, 60, 60)      ' msoShapeOval
    sh.Name = "Oval"
    Set sh = s.Shapes.AddTextbox(1, 10, 100, 120, 40)   ' msoTextOrientationHorizontal
    sh.Name = "Box"
    sh.TextFrame2.TextRange.Text = "Two" & Chr(10) & "lines"
    Set sh = s.Shapes.AddLine(10, 200, 200, 260)
    sh.Name = "Line1"
    Set sh = s.Shapes.AddFormControl(0, 300, 120, 90, 30)   ' xlButtonControl
    sh.Name = "Button1"
    sh.TextFrame.Characters.Text = "Press"
    sh.OnAction = "Clicked"
    Set sh = s.Shapes.AddFormControl(1, 300, 170, 90, 20)   ' xlCheckBox
    sh.Name = "Check1"
    sh.ControlFormat.LinkedCell = "$H$1"
    Set sh = s.Shapes.AddFormControl(2, 300, 200, 90, 20)   ' xlDropDown
    sh.Name = "Drop1"
    sh.ControlFormat.ListFillRange = "$J$1:$J$3"
    Set sh = s.Shapes.AddShape(1, 420, 20, 40, 40)
    sh.Name = "Grouped1"
    Set sh = s.Shapes.AddShape(9, 420, 70, 40, 40)
    sh.Name = "Grouped2"
    s.Shapes.Range(Array("Grouped1", "Grouped2")).Group.Name = "Group1"
    s.Range("J1").Value = "a"
    s.Range("J2").Value = "b"
End Sub

Public Sub Clicked()
End Sub
"""

#: What Excel then says about each one.  Anything this reads is
#: something the reader has to be able to answer the same way.
EXCEL_REPORT = """
Public Function Report() As String
    Dim s As Object, sh As Object, out As String, macro As String
    Set s = ActiveWorkbook.Worksheets(1)
    For Each sh In s.Shapes
        macro = ""
        On Error Resume Next
        macro = sh.OnAction
        On Error GoTo 0
        out = out & sh.Name & vbTab & sh.Type & vbTab & _
            CStr(Round(sh.Left, 2)) & vbTab & CStr(Round(sh.Top, 2)) & vbTab & _
            CStr(Round(sh.Width, 2)) & vbTab & CStr(Round(sh.Height, 2)) & vbTab & _
            macro & vbTab & ShapeText(sh) & vbTab & _
            sh.TopLeftCell.Address & vbTab & sh.BottomRightCell.Address & vbLf
    Next sh
    Report = out
End Function

Private Function ShapeText(sh As Object) As String
    Dim t As String
    On Error Resume Next
    t = sh.TextFrame.Characters.Text
    On Error GoTo 0
    ShapeText = Replace(Replace(t, vbCr, "\\n"), vbLf, "\\n")
End Function
"""

#: PowerPoint: a slide's own shapes, and the macro is an action setting
#: rather than a property of the shape.
POWERPOINT_BUILD = """
Sub Build()
    Dim p As Object, sl As Object, sh As Object
    Set p = ActivePresentation
    Set sl = p.Slides(1)
    Set sh = sl.Shapes.AddShape(1, 20, 30, 120, 60)
    sh.Name = "Rect"
    sh.TextFrame.TextRange.Text = "Hello"
    sh.ActionSettings(1).Action = 8            ' ppActionRunMacro
    sh.ActionSettings(1).Run = "Clicked"
    Set sh = sl.Shapes.AddShape(9, 200, 30, 80, 80)
    sh.Name = "Oval"
    Set sh = sl.Shapes.AddTextbox(1, 20, 120, 160, 40)
    sh.Name = "Box"
    sh.TextFrame.TextRange.Text = "Two" & Chr(13) & "lines"
    Set sh = sl.Shapes.AddLine(20, 200, 200, 240)
    sh.Name = "Line1"
    Set sh = sl.Shapes.AddShape(1, 320, 30, 40, 40)
    sh.Name = "Grouped1"
    Set sh = sl.Shapes.AddShape(9, 320, 80, 40, 40)
    sh.Name = "Grouped2"
    sl.Shapes.Range(Array("Grouped1", "Grouped2")).Group.Name = "Group1"
End Sub

Public Sub Clicked()
End Sub
"""

POWERPOINT_REPORT = """
Public Function Report() As String
    Dim sl As Object, sh As Object, out As String, macro As String
    Set sl = ActivePresentation.Slides(1)
    For Each sh In sl.Shapes
        macro = ""
        On Error Resume Next
        If sh.ActionSettings(1).Action = 8 Then macro = sh.ActionSettings(1).Run
        On Error GoTo 0
        out = out & sh.Name & vbTab & sh.Type & vbTab & _
            CStr(Round(sh.Left, 2)) & vbTab & CStr(Round(sh.Top, 2)) & vbTab & _
            CStr(Round(sh.Width, 2)) & vbTab & CStr(Round(sh.Height, 2)) & vbTab & _
            macro & vbTab & ShapeText(sh) & vbLf
    Next sh
    Report = out
End Function

Private Function ShapeText(sh As Object) As String
    Dim t As String
    On Error Resume Next
    t = sh.TextFrame.TextRange.Text
    On Error GoTo 0
    ShapeText = Replace(Replace(t, vbCr, "\\n"), vbLf, "\\n")
End Function
"""

#: Word: a shape is either in the text or anchored to it, and Word has
#: no macro to run from one.
WORD_BUILD = """
Sub Build()
    Dim d As Object, sh As Object
    Set d = ActiveDocument
    d.Content.Text = "First paragraph." & vbCr & "Second paragraph." & vbCr
    Set sh = d.Shapes.AddShape(1, 40, 60, 120, 50)
    sh.Name = "Rect"
    sh.TextFrame.TextRange.Text = "Hello"
    Set sh = d.Shapes.AddShape(9, 200, 60, 70, 70)
    sh.Name = "Oval"
    Set sh = d.Shapes.AddTextbox(1, 40, 150, 140, 40)
    sh.Name = "Box"
    sh.TextFrame.TextRange.Text = "Two" & Chr(11) & "lines"
    Set sh = d.Shapes.AddLine(40, 230, 220, 260)
    sh.Name = "Line1"
    ' No group here: Word moves the anchor of the first shape, so two
    ' shapes added this way never share one, and Group refuses. A Word
    ' group is made in a drawing canvas, which is its own fixture.
    Set sh = d.Shapes.AddShape(1, 300, 60, 40, 40)
    sh.Name = "Plain1"
    d.InlineShapes.AddHorizontalLineStandard d.Paragraphs(1).Range
End Sub
"""

WORD_REPORT = """
Public Function Report() As String
    Dim d As Object, sh As Object, out As String, i As Long
    Set d = ActiveDocument
    For Each sh In d.Shapes
        out = out & sh.Name & vbTab & sh.Type & vbTab & _
            CStr(Round(sh.Left, 2)) & vbTab & CStr(Round(sh.Top, 2)) & vbTab & _
            CStr(Round(sh.Width, 2)) & vbTab & CStr(Round(sh.Height, 2)) & vbTab & _
            "anchored" & vbTab & ShapeText(sh) & vbLf
    Next sh
    For i = 1 To d.InlineShapes.Count
        out = out & "inline" & CStr(i) & vbTab & d.InlineShapes(i).Type & vbTab & _
            "" & vbTab & "" & vbTab & _
            CStr(Round(d.InlineShapes(i).Width, 2)) & vbTab & _
            CStr(Round(d.InlineShapes(i).Height, 2)) & vbTab & _
            "inline" & vbTab & "" & vbLf
    Next i
    Report = out
End Function

Private Function ShapeText(sh As Object) As String
    Dim t As String
    On Error Resume Next
    t = sh.TextFrame.TextRange.Text
    On Error GoTo 0
    ShapeText = Replace(Replace(Replace(t, vbCr, "\\n"), vbLf, "\\n"), Chr(11), "\\n")
End Function
"""

#: What each host's report columns mean, in order.
COLUMNS = {
    "excel": ("name", "type", "left", "top", "width", "height", "macro", "text", "from", "to"),
    "powerpoint": ("name", "type", "left", "top", "width", "height", "macro", "text"),
    "word": ("name", "type", "left", "top", "width", "height", "placement", "text"),
}


def _rows(body: str, host: str) -> list[dict[str, str]]:
    out: list[dict[str, str]] = []
    for line in body.split("\n"):
        if not line.strip():
            continue
        out.append(dict(zip(COLUMNS[host], line.split("\t"))))
    return out


def measure_excel() -> int:
    from pyvbaharness import ExcelSession

    from pyopenvba.excel import ExcelFile

    FOLDER.mkdir(parents=True, exist_ok=True)
    target = FOLDER / "excel_shapes.xlsm"
    seed = FOLDER / "_seed.xlsm"
    ExcelFile.create_new(seed)
    with ExcelSession() as excel:
        excel.open_document(seed)
        result = excel.run_vba(
            EXCEL_BUILD + EXCEL_REPORT + _saver("excel", target),
            "BuildAndReport",
            timeout=300.0,
        )
        if not result.ok:
            print(f"failed: {result.outcome} {result.message}")
            return 1
        body = str(result.value or "")
    seed.unlink(missing_ok=True)
    return _write("excel", body, target)


def measure_powerpoint() -> int:
    from pyvbaharness import PowerPointSession

    from pyopenvba.powerpoint import PowerPointFile

    FOLDER.mkdir(parents=True, exist_ok=True)
    target = FOLDER / "powerpoint_shapes.pptm"
    seed = FOLDER / "_seed.pptm"
    PowerPointFile.create_new(seed)
    with PowerPointSession() as ppt:
        ppt.open_document(seed)
        result = ppt.run_vba(
            POWERPOINT_BUILD + POWERPOINT_REPORT + _saver("powerpoint", target),
            "BuildAndReport",
            timeout=300.0,
        )
        if not result.ok:
            print(f"failed: {result.outcome} {result.message}")
            return 1
        body = str(result.value or "")
    seed.unlink(missing_ok=True)
    return _write("powerpoint", body, target)


def measure_word() -> int:
    from pyvbaharness import WordSession

    from pyopenvba.word import WordFile

    FOLDER.mkdir(parents=True, exist_ok=True)
    target = FOLDER / "word_shapes.docm"
    seed = FOLDER / "_seed.docm"
    WordFile.create_new(seed)
    with WordSession() as word:
        word.open_document(seed)
        result = word.run_vba(
            WORD_BUILD + WORD_REPORT + _saver("word", target),
            "BuildAndReport",
            timeout=300.0,
        )
        if not result.ok:
            print(f"failed: {result.outcome} {result.message}")
            return 1
        body = str(result.value or "")
    seed.unlink(missing_ok=True)
    return _write("word", body, target)


#: How each host is told to save, with its own file format number: 52
#: is xlsm, 25 is pptm, 13 is docm.
SAVE_AS = {
    "excel": 'ActiveWorkbook.SaveAs "{path}", 52',
    "powerpoint": 'ActivePresentation.SaveAs "{path}", 25',
    "word": 'ActiveDocument.SaveAs2 "{path}", 13',
}


def _saver(host: str, target: Path) -> str:
    """Build, report, then save where the fixture goes.

    One procedure rather than three calls: the session opens the file
    once, and the save has to happen after the shapes are there.
    """
    return (
        "Public Function BuildAndReport() As String\n"
        "    Build\n"
        "    " + SAVE_AS[host].format(path=target) + "\n"
        "    BuildAndReport = Report\n"
        "End Function\n"
    )


def _write(host: str, body: str, target: Path) -> int:
    rows = _rows(body, host)
    if not rows:
        print("no shapes came back")
        return 1
    out = FOLDER / f"{host}_shapes.json"
    out.write_text(
        json.dumps({"host": host, "file": target.name, "shapes": rows}, indent=2) + "\n",
        encoding="utf-8",
    )
    print(f"{len(rows)} shapes -> {out.relative_to(ROOT)}, {target.relative_to(ROOT)}")
    return 0


def main() -> int:
    which = sys.argv[1] if len(sys.argv) > 1 else "all"
    hosts = {"excel": measure_excel, "powerpoint": measure_powerpoint, "word": measure_word}
    if which == "all":
        return max(run() for run in hosts.values())
    if which not in hosts:
        print(f"usage: measure_shapes.py [{' | '.join(hosts)} | all]")
        return 2
    return hosts[which]()


if __name__ == "__main__":
    sys.exit(main())
