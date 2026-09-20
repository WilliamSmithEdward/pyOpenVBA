"""Live Excel gate for the in-memory workbook (opt-in).

Two things only real Excel can answer.  First, that a file this wrote
opens without a repair prompt and holds what the macro put in it: the
offline tests prove the bytes are what we meant, not that Excel agrees.
Second, that a macro run here and the same macro run in Excel leave the
same cells behind.

Opt-in: set ``RUN_LIVE_EXCEL=1`` on a Windows machine with desktop Excel
installed.  Skipped everywhere else, including CI.
"""

from __future__ import annotations

import os
import shutil
import sys
from pathlib import Path

import pytest

from pyopenvba.apps.excel import ExcelApplication

pytestmark = pytest.mark.skipif(
    os.environ.get("RUN_LIVE_EXCEL") != "1" or sys.platform != "win32",
    reason="live Excel gate: set RUN_LIVE_EXCEL=1 on Windows with Excel installed",
)

FIXTURES = Path(__file__).parent / "fixtures" / "power_query"

#: The macro both sides run.  Everything it touches is something the
#: model claims to implement.
MACRO = """
Sub Fill()
    Dim r As Long
    Range("A1").Value = "Item"
    Range("B1").Value = "Qty"
    For r = 1 To 5
        Cells(r + 1, 1).Value = "Item " & r
        Cells(r + 1, 2).Value = r * 10
    Next r
    Range("D1").Value = Application.WorksheetFunction.Sum(Range("B2:B6"))
    Range("D2").Value = Range("A1").End(xlDown).Address
    Range("D3").Value = ActiveSheet.UsedRange.Address
    Range("D4").Value = Range("B2:B6").Count
End Sub
"""


def _excel_says(source: str, macro: str) -> dict[str, object]:
    """Run the macro in real Excel and read the cells back."""
    harness = pytest.importorskip("pyvbaharness")

    reader = (
        "Public Function Report() As String\n"
        "    Fill\n"
        '    Report = Range("D1").Value & "|" & Range("D2").Value & "|" & _\n'
        '        Range("D3").Value & "|" & Range("D4").Value & "|" & _\n'
        '        Range("A6").Value & "|" & TypeName(Range("B2").Value)\n'
        "End Function\n"
    )
    with harness.ExcelSession() as excel:
        excel.new_document()
        result = excel.run_vba(macro + reader, "Report", timeout=120.0)
        assert result.ok, f"{result.outcome}: {result.message}"
        pieces = str(result.value).split("|")
    return dict(zip(("total", "end", "used", "count", "a6", "type"), pieces))


def _model_says(macro: str) -> dict[str, object]:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(macro, name="Module1")
    app.run("Fill")
    sheet = app.sheet(1)
    return {
        "total": app.evaluate('CStr(Range("D1").Value)'),
        "end": sheet.value("D2"),
        "used": sheet.value("D3"),
        "count": app.evaluate('CStr(Range("D4").Value)'),
        "a6": sheet.value("A6"),
        "type": app.evaluate('TypeName(Range("B2").Value)'),
    }


def test_the_model_and_excel_agree_on_what_the_macro_did() -> None:
    assert _model_says(MACRO) == _excel_says("", MACRO)


def test_a_workbook_this_wrote_opens_in_excel(tmp_path: Path) -> None:
    harness = pytest.importorskip("pyvbaharness")

    app = ExcelApplication()
    app.add_workbook()
    app.add_module(MACRO, name="Module1")
    app.run("Fill")
    out = tmp_path / "written.xlsx"
    app.save(out)

    with harness.ExcelSession() as excel:
        excel.open_document(out)
        result = excel.run_vba(
            "Public Function Report() As String\n"
            '    Report = ActiveSheet.Range("A1").Value & "|" & _\n'
            '        CStr(ActiveSheet.Range("B6").Value) & "|" & _\n'
            '        ActiveSheet.UsedRange.Address\n'
            "End Function\n",
            "Report",
            timeout=120.0,
        )
        assert result.ok, f"{result.outcome}: {result.message}"
        assert str(result.value) == "Item|50|$A$1:$D$6"


#: A workbook of formulas, built here and recalculated by Excel.
FORMULAS = """
Sub Build()
    Dim r As Long
    For r = 1 To 5
        Cells(r, 1).Value = r
        Cells(r, 2).Value = r * 1.5
    Next r
    Range("C1:C5").Formula = "=A1*B1"
    Range("E1").Formula = "=SUM(C1:C5)"
    Range("E2").Formula = "=ROUND(AVERAGE(B1:B5),2)"
    Range("E3").Formula = "=IF(E1>20,""over"",""under"")"
    Range("E4").Formula = "=COUNTIF(A1:A5,"">2"")"
    Range("E5").Formula = "=VLOOKUP(3,A1:C5,3,FALSE)"
    Range("E6").Formula = "=TEXT(E1,""#,##0.00"")"
    Range("E7").Formula = "=A1&""-""&B1"
    Range("E8").Formula = "=IFERROR(1/0,""caught"")"
    Range("E9").Formula = "=MOD(-7,3)"
    Range("E10").Formula = "=INDEX(C1:C5,MATCH(4,A1:A5,0))"
End Sub
"""

#: Where the answers are read from, on both sides.
ANSWERS = [f"E{number}" for number in range(1, 11)] + ["C1", "C5"]


def test_excel_agrees_with_what_the_engine_calculated(tmp_path: Path) -> None:
    """The same formulas, worked out here and by Excel, in a file it opened.

    This is the check the offline probes cannot make: that the
    calculated values survive the save, that Excel accepts the file, and
    that recalculating there lands on the same numbers.
    """
    harness = pytest.importorskip("pyvbaharness")

    app = ExcelApplication()
    app.add_workbook()
    app.add_module(FORMULAS, name="Module1")
    app.run("Build")
    ours = [str(app.evaluate(f'CStr(Range("{where}").Value)')) for where in ANSWERS]

    out = tmp_path / "formulas.xlsx"
    app.save(out)

    reader = (
        "Public Function Report() As String\n"
        "    Application.CalculateFull\n"
        "    Dim out As String, i As Long\n"
        "    Dim cells As Variant\n"
        f"    cells = Array({', '.join(repr(where).replace(chr(39), chr(34)) for where in ANSWERS)})\n"
        "    For i = LBound(cells) To UBound(cells)\n"
        '        out = out & CStr(ActiveSheet.Range(cells(i)).Value) & "|"\n'
        "    Next i\n"
        "    Report = out\n"
        "End Function\n"
    )
    with harness.ExcelSession() as excel:
        excel.open_document(out)
        result = excel.run_vba(reader, "Report", timeout=120.0)
        assert result.ok, f"{result.outcome}: {result.message}"
        theirs = str(result.value).rstrip("|").split("|")

    assert theirs == ours


def test_an_edited_workbook_still_opens_with_its_query(tmp_path: Path) -> None:
    """A Power Query workbook edited here keeps what Excel needs.

    The model has no field for a connection or a query table, so this is
    the check that not rewriting them was enough.
    """
    harness = pytest.importorskip("pyvbaharness")

    copy = tmp_path / "edited.xlsx"
    shutil.copyfile(FIXTURES / "loaded_to_sheet.xlsx", copy)
    app = ExcelApplication.open(copy)
    app.add_module('Sub Touch()\n    Worksheets(1).Range("E1").Value = 99\nEnd Sub\n', name="Module1")
    app.run("Touch")
    out = tmp_path / "edited_out.xlsx"
    app.save(out)

    with harness.ExcelSession() as excel:
        excel.open_document(out)
        result = excel.run_vba(
            "Public Function Report() As String\n"
            '    Report = CStr(ActiveSheet.Range("E1").Value) & "|" & _\n'
            "        CStr(ActiveWorkbook.Queries.Count) & \"|\" & _\n"
            "        CStr(ActiveWorkbook.Connections.Count)\n"
            "End Function\n",
            "Report",
            timeout=120.0,
        )
        assert result.ok, f"{result.outcome}: {result.message}"
        assert str(result.value) == "99|1|1"


#: The shapes a macro makes here, which Excel then has to report back
#: the same way.
SHAPES = """
Sub Draw()
    Dim sh As Object
    Set sh = ActiveSheet.Shapes.AddShape(1, 10, 20, 100, 50)
    sh.Name = "Rect"
    sh.TextFrame.Characters.Text = "Hello"
    sh.OnAction = "Clicked"
    Set sh = ActiveSheet.Shapes.AddShape(9, 150, 30, 60, 60)
    sh.Name = "Oval"
    Set sh = ActiveSheet.Shapes.AddTextbox(1, 10, 120, 120, 40)
    sh.Name = "Box"
    sh.TextFrame.Characters.Text = "Two" & Chr(10) & "lines"
End Sub

Public Sub Clicked()
End Sub
"""


def test_excel_opens_the_shapes_this_wrote(tmp_path: Path) -> None:
    """Shapes made by a macro here, read back by Excel from the file.

    The offline tests prove the drawing markup is what we meant; only
    Excel can say whether it opens the workbook and sees the same
    shapes in the same places, with the macro still on the first one.
    """
    harness = pytest.importorskip("pyvbaharness")

    app = ExcelApplication()
    app.add_workbook()
    app.add_module(SHAPES, name="Module1")
    app.run("Draw")
    ours = [
        str(app.evaluate(f"CStr(ActiveSheet.Shapes({number}).{what})"))
        for number in (1, 2, 3)
        for what in ("Name", "Type", "Left", "Top", "Width", "Height", "OnAction")
    ]
    text = str(app.evaluate('ActiveSheet.Shapes("Box").TextFrame.Characters.Text'))
    ours.append(text.replace(chr(10), "/"))

    out = tmp_path / "shapes.xlsm"
    app.save(out)

    reader = (
        "Public Function Report() As String\n"
        "    Dim out As String, i As Long, what As Variant, one As Variant\n"
        '    what = Array("Name", "Type", "Left", "Top", "Width", "Height", "OnAction")\n'
        "    For i = 1 To 3\n"
        "        For Each one In what\n"
        "            out = out & CStr(CallByName(ActiveSheet.Shapes(i), CStr(one), VbGet)) & vbLf\n"
        "        Next one\n"
        "    Next i\n"
        '    Dim body As String\n'
        '    body = ActiveSheet.Shapes("Box").TextFrame.Characters.Text\n'
        '    out = out & Replace(Replace(body, vbCr, "/"), vbLf, "/") & vbLf\n'
        "    Report = out\n"
        "End Function\n"
    )
    with harness.ExcelSession() as excel:
        excel.open_document(out)
        result = excel.run_vba(reader, "Report", timeout=120.0)
        assert result.ok, f"{result.outcome}: {result.message}"
        theirs = str(result.value).rstrip("\n").split("\n")

    assert [one.replace("\r", "") for one in theirs] == ours


#: A form control, which is four parts of a workbook at once.
CONTROL = """
Sub Draw()
    Dim sh As Object
    Set sh = ActiveSheet.Shapes.AddFormControl(0, 100, 50, 90, 30)
    sh.Name = "Go"
    sh.OnAction = "Clicked"
    sh.TextFrame.Characters.Text = "Press me"
End Sub

Public Sub Clicked()
End Sub
"""


def test_excel_opens_a_form_control_this_made(tmp_path: Path) -> None:
    """A button built from nothing, and Excel agrees it is one.

    A control is the drawing, the sheet's own record of it, a part
    saying what kind it is and the VML Excel draws it from.  Any of
    them wrong and Excel refuses the file rather than saying why, so
    this gate is the only way to know.
    """
    harness = pytest.importorskip("pyvbaharness")

    app = ExcelApplication()
    app.add_workbook()
    app.add_module(CONTROL, name="Module1")
    app.run("Draw")
    out = tmp_path / "control.xlsm"
    app.save(out)

    reader = (
        "Public Function Report() As String\n"
        "    Dim sh As Object\n"
        "    Set sh = ActiveSheet.Shapes(1)\n"
        '    Report = sh.Name & "|" & CStr(sh.Type) & "|" & sh.OnAction & "|" & _\n'
        '        CStr(sh.Left) & "|" & CStr(sh.Width) & "|" & sh.TextFrame.Characters.Text & _\n'
        '        "|" & CStr(sh.FormControlType)\n'
        "End Function\n"
    )
    with harness.ExcelSession() as excel:
        excel.open_document(out)
        result = excel.run_vba(reader, "Report", timeout=120.0)
        assert result.ok, f"{result.outcome}: {result.message}"
        assert str(result.value) == "Go|8|control.xlsm!Clicked|100|90|Press me|0"


#: A query that does enough to be worth comparing: a group, a sort, a
#: derived column and a rounding, all of which the M tests measure one
#: at a time.
QUERY = """let
    Source = #table({"Region", "Amount"}, {{"West", 10}, {"East", 20}, {"West", 5}}),
    Grouped = Table.Group(Source, {"Region"}, {{"Total", each List.Sum([Amount])}}),
    Sorted = Table.Sort(Grouped, {{"Region", Order.Ascending}}),
    Final = Table.AddColumn(Sorted, "Share", each Number.Round([Total] / 35, 3))
in
    Final"""

#: The block the query fills, headers included.
LANDED = ("A1", "B1", "C1", "A2", "B2", "C2", "A3", "B3", "C3")


def test_the_refresh_lands_what_power_query_lands(tmp_path: Path) -> None:
    """The same query, refreshed here and by Excel, fills the same cells.

    The probe file holds the evaluator to Power Query expression by
    expression.  This is the other half: that a whole query, loaded to a
    sheet, puts the same values in the same places as the engine does.
    """
    harness = pytest.importorskip("pyvbaharness")

    from pyopenvba import PowerQueryWorkbook
    from pyopenvba.excel import ExcelFile

    seed = tmp_path / "seed.xlsm"
    source = tmp_path / "query.xlsm"
    ExcelFile.create_new(seed)
    book = PowerQueryWorkbook(seed)
    book.add_query("Summary", QUERY)
    book.load_to_sheet("Summary", ["Region", "Total", "Share"])
    book.save(source)

    app = ExcelApplication.open(source)
    app.refresh_query("Summary")
    ours = [str(app.evaluate(f'CStr(Worksheets(1).Range("{where}").Value)')) for where in LANDED]

    reader = (
        "Public Function Report() As String\n"
        "    Dim c As WorkbookConnection\n"
        "    For Each c In ActiveWorkbook.Connections\n"
        "        On Error Resume Next\n"
        "        c.OLEDBConnection.BackgroundQuery = False\n"
        "        On Error GoTo 0\n"
        "    Next c\n"
        "    ActiveWorkbook.RefreshAll\n"
        "    Application.CalculateUntilAsyncQueriesDone\n"
        "    Dim out As String, i As Long, cells As Variant\n"
        f"    cells = Array({', '.join(chr(34) + where + chr(34) for where in LANDED)})\n"
        "    For i = LBound(cells) To UBound(cells)\n"
        '        out = out & CStr(Worksheets(1).Range(cells(i)).Value) & "|"\n'
        "    Next i\n"
        "    Report = out\n"
        "End Function\n"
    )
    with harness.ExcelSession() as excel:
        excel.open_document(source)
        result = excel.run_vba(reader, "Report", timeout=300.0)
        assert result.ok, f"{result.outcome}: {result.message}"
        theirs = str(result.value).rstrip("|").split("|")

    assert ours == theirs
