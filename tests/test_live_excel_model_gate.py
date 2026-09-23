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
        '    Report = ActiveSheet.StandardHeight & "#" & sh.Name & "|" & CStr(sh.Type) & "|" & sh.OnAction & "|" & _\n'
        '        CStr(sh.Left) & "|" & CStr(sh.Width) & "|" & sh.TextFrame.Characters.Text & _\n'
        '        "|" & CStr(sh.FormControlType)\n'
        "End Function\n"
    )
    with harness.ExcelSession() as excel:
        excel.open_document(out)
        result = excel.run_vba(reader, "Report", timeout=120.0)
        assert result.ok, f"{result.outcome}: {result.message}"
    standard, answer = str(result.value).split("#", 1)
    left, width = _on_display(standard, 100, 90)
    assert answer == f"Go|8|control.xlsm!Clicked|{left}|{width}|Press me|0"


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


def test_excel_opens_public_python_shape_edits(tmp_path: Path) -> None:
    """Creation, edits and control removal through Python agree with live Excel."""
    harness = pytest.importorskip("pyvbaharness")
    source = Path(__file__).parent / "fixtures" / "shapes" / "excel_shapes.xlsm"
    app = ExcelApplication.open(source, with_vba=False)
    sheet = app.sheet(1)
    sheet.update_shape("Button1", new_name="Run", left=40, top=60, width=110, height=35,
                       text="Changed\nCaption", macro="")
    sheet.remove_shape("Drop1")
    sheet.add_button(name="New", left=300, top=220, text="New button", macro="Clicked")
    out = tmp_path / "python_shapes.xlsm"
    app.save(out)
    reader = (
        "Public Function Report() As String\n"
        "    Dim sh As Object\n"
        '    Set sh = ActiveSheet.Shapes("Run")\n'
        '    Report = ActiveSheet.StandardHeight & "#" & CStr(ActiveSheet.Shapes.Count) & "|" & sh.Name & "|" & _\n'
        '        CStr(sh.Left) & "|" & CStr(sh.Top) & "|" & CStr(sh.Width) & "|" & _\n'
        '        CStr(sh.Height) & "|" & Replace(Replace(sh.TextFrame.Characters.Text, vbCr, ""), vbLf, "/") & "|" & _\n'
        '        sh.OnAction & "|" & ActiveSheet.Shapes("New").TextFrame.Characters.Text & "|" & _\n'
        '        CStr(ActiveSheet.Shapes("Check1").Type)\n'
        "End Function\n"
    )
    with harness.ExcelSession() as excel:
        excel.open_document(out)
        result = excel.run_vba(reader, "Report", timeout=120.0)
        assert result.ok, f"{result.outcome}: {result.message}"
    standard, answer = str(result.value).split("#", 1)
    left, width = _on_display(standard, 40, 110)
    top, height = _on_display(standard, 60, 35)
    assert answer == f"9|Run|{left}|{top}|{width}|{height}|Changed/Caption||New button|8"


def _on_display(standard_height: str, start: float, size: float) -> tuple[str, str]:
    """Where Excel puts a shape's edges once it loads them from the anchor, as CStr spells them.

    Excel rounds each edge to a whole pixel of its display: 0.75pt at
    96 DPI, where a standard row is 15pt, and 0.5pt at 144 DPI, where it
    is 14.5pt. So a button written at 40pt reads 39.75 on the first and
    40 on the second.
    """
    pixel = {"15": 0.75, "14.5": 0.5}.get(standard_height)
    if pixel is None:
        pytest.skip(f"no pixel grid measured for a display whose standard row is {standard_height}pt")
    first = round(start / pixel) * pixel
    last = round((start + size) / pixel) * pixel

    def spelled(value: float) -> str:
        return f"{value:.15g}"

    return spelled(first), spelled(last - first)


def test_excel_opens_after_last_public_button_is_deleted(tmp_path: Path) -> None:
    harness = pytest.importorskip("pyvbaharness")
    app = ExcelApplication()
    app.add_workbook()
    sheet = app.sheet(1)
    sheet.add_button(name="Gone", text="Delete me")
    out = tmp_path / "deleted_button.xlsm"
    app.save(out)
    sheet.remove_shape("Gone")
    app.save(out)
    with harness.ExcelSession() as excel:
        excel.open_document(out)
        result = excel.run_vba(
            "Public Function Report() As String\n"
            "    Report = CStr(ActiveSheet.Shapes.Count)\nEnd Function\n",
            "Report", timeout=120.0,
        )
        assert result.ok, f"{result.outcome}: {result.message}"
        assert str(result.value) == "0"


def test_excel_vba_default_shape_name_can_duplicate_a_manual_name() -> None:
    harness = pytest.importorskip("pyvbaharness")
    app = ExcelApplication()
    app.add_workbook()
    source = (
        "Public Function Report() As String\n"
        '    ActiveSheet.Shapes.AddShape(1, 0, 0, 100, 50).Name = "Rectangle 2"\n'
        "    Report = ActiveSheet.Shapes.AddShape(1, 0, 0, 100, 50).Name\n"
        "End Function\n"
    )
    app.add_module(source, name="NameProbe")
    ours = app.run("Report")
    with harness.ExcelSession() as excel:
        excel.new_document()
        result = excel.run_vba(
            source, "Report", timeout=120.0,
        )
        assert result.ok, f"{result.outcome}: {result.message}"
        assert str(result.value) == ours


@pytest.mark.parametrize("clear", [False, True])
def test_excel_reads_python_control_bindings(tmp_path: Path, clear: bool) -> None:
    harness = pytest.importorskip("pyvbaharness")
    source = Path(__file__).parent / "fixtures" / "shapes" / "excel_shapes.xlsm"
    app = ExcelApplication.open(source, with_vba=False)
    sheet = app.sheet(1)
    sheet.update_control("Check1", linked_cell="" if clear else "$K$2")
    sheet.update_control("Drop1", linked_cell="" if clear else "$K$3",
                         list_range="" if clear else "$J$2:$J$3")
    out = tmp_path / "bindings.xlsm"
    app.save(out)
    with harness.ExcelSession() as excel:
        excel.open_document(out)
        result = excel.run_vba(
            'Public Function Report() As String\n'
            'Report = ActiveSheet.Shapes("Check1").ControlFormat.LinkedCell & "|" & _\n'
            'ActiveSheet.Shapes("Drop1").ControlFormat.LinkedCell & "|" & _\n'
            'ActiveSheet.Shapes("Drop1").ControlFormat.ListFillRange\nEnd Function\n',
            "Report", timeout=120.0,
        )
        assert result.ok, f"{result.outcome}: {result.message}"
        assert str(result.value) == ("||" if clear else "$K$2|$K$3|$J$2:$J$3")


@pytest.mark.parametrize("value", [-4146, 1, 2])
def test_excel_reads_python_checkbox_values(tmp_path: Path, value: int) -> None:
    harness = pytest.importorskip("pyvbaharness")
    source = Path(__file__).parent / "fixtures" / "shapes" / "excel_shapes.xlsm"
    app = ExcelApplication.open(source, with_vba=False)
    sheet = app.sheet(1)
    sheet.set_control_value("Check1", 1)
    sheet.set_control_value("Check1", value)
    out = tmp_path / "checkbox.xlsm"
    app.save(out)
    with harness.ExcelSession(harness.HarnessConfig(lock_wait_s=45.0)) as excel:
        excel.open_document(out)
        result = excel.run_vba(
            'Public Function Report() As String\n'
            'Report = CStr(ActiveSheet.Shapes("Check1").ControlFormat.Value) & "|" & _\n'
            'CStr(Range("H1").Value)\nEnd Function\n', "Report", timeout=120.0,
        )
        assert result.ok, f"{result.outcome}: {result.message}"
        expected = "Error 2042" if value == 2 else str(value == 1)
        assert str(result.value) == f"{value}|{expected}"


def test_excel_reads_checkbox_created_by_headless_vba(tmp_path: Path) -> None:
    harness = pytest.importorskip("pyvbaharness")
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('''Public Sub Build()
Dim sh As Object
Set sh = ActiveSheet.Shapes.AddFormControl(1, 0, 0, 90, 20)
sh.Name = "Created"
sh.ControlFormat.LinkedCell = "$H$1"
sh.ControlFormat.Value = 2
End Sub''', name="Builder")
    app.run("Build")
    out = tmp_path / "created_checkbox.xlsm"
    app.save(out)
    with harness.ExcelSession() as excel:
        excel.open_document(out)
        result = excel.run_vba(
            'Public Function Report() As String\n'
            'Report = CStr(ActiveSheet.Shapes("Created").FormControlType) & "|" & _\n'
            'CStr(ActiveSheet.Shapes("Created").ControlFormat.Value) & "|" & _\n'
            'CStr(Range("H1").Value)\nEnd Function\n', "Report", timeout=120.0,
        )
        assert result.ok, f"{result.outcome}: {result.message}"
        assert str(result.value) == "1|2|Error 2042"


@pytest.mark.parametrize("kind", [2, 6])
def test_excel_reads_headless_list_selection(tmp_path: Path, kind: int) -> None:
    harness = pytest.importorskip("pyvbaharness")
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(f'''Public Sub Build()
Dim sh As Object
Range("J1").Value = "a"
Range("J2").Value = "b"
Range("J3").Value = "c"
Set sh = ActiveSheet.Shapes.AddFormControl({kind}, 0, 0, 90, 60)
sh.Name = "Choices"
sh.ControlFormat.ListFillRange = "$J$1:$J$3"
sh.ControlFormat.LinkedCell = "$H$1"
sh.ControlFormat.Value = 2
End Sub''', name="Builder")
    app.run("Build")
    out = tmp_path / "choices.xlsm"
    app.save(out)
    # Reopen and edit, to exercise existing-part persistence too.
    app = ExcelApplication.open(out, with_vba=False)
    app.sheet(1).set_control_value("Choices", 3)
    app.save(out)
    with harness.ExcelSession() as excel:
        excel.open_document(out)
        result = excel.run_vba(
            'Public Function Report() As String\n'
            'Report = CStr(ActiveSheet.Shapes("Choices").FormControlType) & "|" & _\n'
            'CStr(ActiveSheet.Shapes("Choices").ControlFormat.Value) & "|" & _\n'
            'CStr(ActiveSheet.Shapes("Choices").ControlFormat.ListCount) & "|" & _\n'
            'CStr(Range("H1").Value)\nEnd Function\n', "Report", timeout=120.0,
        )
        assert result.ok, f"{result.outcome}: {result.message}"
        assert str(result.value) == f"{kind}|3|3|3"


@pytest.mark.parametrize("mode", [1, 2, 3])
def test_excel_reads_inline_items_and_multi_selection(tmp_path: Path, mode: int) -> None:
    harness = pytest.importorskip("pyvbaharness")
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('''Public Sub Build()
Dim sh As Object
Set sh = ActiveSheet.Shapes.AddFormControl(6, 0, 0, 90, 60)
sh.Name = "Choices"
End Sub''', name="Builder")
    app.run("Build")
    sheet = app.sheet(1)
    for text in ("one", "two & three", "four"):
        sheet.add_control_item("Choices", text)
    sheet.set_control_selection_mode("Choices", mode)
    if mode == 1:
        sheet.set_control_value("Choices", 2)
    else:
        sheet.set_control_selection("Choices", [1, 3])
    out = tmp_path / "inline.xlsm"
    app.save(out)
    app = ExcelApplication.open(out, with_vba=False)
    app.sheet(1).update_control_item("Choices", 2, "two & edited")
    app.save(out)
    with harness.ExcelSession(harness.HarnessConfig(lock_wait_s=45.0)) as excel:
        excel.open_document(out)
        result = excel.run_vba(
            'Public Function Report() As String\nDim sh As Object\n'
            'Set sh = ActiveSheet.Shapes("Choices")\n'
            'Report = CStr(sh.ControlFormat.MultiSelect) & "|" & sh.ControlFormat.List(2) & "|" & _\n'
            'CStr(sh.DrawingObject.Selected(1)) & "|" & CStr(sh.DrawingObject.Selected(2)) & "|" & _\n'
            'CStr(sh.DrawingObject.Selected(3))\nEnd Function\n', "Report", timeout=120.0,
        )
        assert result.ok, f"{result.outcome}: {result.message}"
        flags = "False|True|False" if mode == 1 else "True|False|True"
        native_mode = {1: -4142, 2: -4154, 3: 3}[mode]
        assert str(result.value) == f"{native_mode}|two & edited|{flags}"


@pytest.mark.parametrize("operation", ["replace", "append", "clear", "remove"])
def test_excel_reads_range_backed_list_edits(tmp_path: Path, operation: str) -> None:
    harness = pytest.importorskip("pyvbaharness")
    source = Path(__file__).parent / "fixtures" / "shapes" / "excel_shapes.xlsm"
    app = ExcelApplication.open(source, with_vba=False)
    sheet = app.sheet(1)
    sheet.update_control("Drop1", linked_cell="$H$2")
    sheet.set_control_value("Drop1", 2)
    if operation == "replace":
        app.add_module('''Public Sub ReplaceList()
ActiveSheet.Shapes("Drop1").ControlFormat.List = Array("x", "y")
End Sub''', name="Replacement")
        app.run("ReplaceList")
    elif operation == "append":
        sheet.add_control_item("Drop1", "x")
    elif operation == "clear":
        sheet.clear_control_items("Drop1")
    else:
        with pytest.raises(ValueError):
            sheet.remove_control_item("Drop1", 2)
    out = tmp_path / "range_edit.xlsm"
    app.save(out)
    with harness.ExcelSession(harness.HarnessConfig(lock_wait_s=45.0)) as excel:
        excel.open_document(out)
        result = excel.run_vba(
            'Public Function Report() As String\nDim cf As Object, i As Long\n'
            'Set cf = ActiveSheet.Shapes("Drop1").ControlFormat\n'
            'Report = cf.ListFillRange & "|" & CStr(cf.Value) & "|" & CStr(Range("H2").Value) & "|"\n'
            'For i = 1 To cf.ListCount\nReport = Report & cf.List(i) & ";"\nNext i\n'
            'Report = Report & "|" & Range("J1").Value & ";" & Range("J2").Value\nEnd Function\n',
            "Report", timeout=120.0,
        )
        assert result.ok, f"{result.outcome}: {result.message}"
        expected = {"replace": "|0|0|x;y;|a;b", "append": "|0|0|x;|a;b",
                    "clear": "|0|0||a;b", "remove": "$J$1:$J$3|2|2|a;b;;|a;b"}
        assert str(result.value) == expected[operation]


def test_excel_opens_issue25_control_types_and_defaults(tmp_path: Path) -> None:
    harness = pytest.importorskip("pyvbaharness")
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('''Public Sub Build()
ActiveSheet.Shapes.AddFormControl(7, 0, 0, 90, 30).Name = "Radio"
ActiveSheet.Shapes.AddFormControl(9, 0, 40, 90, 30).Name = "Spin"
ActiveSheet.Shapes.AddFormControl(8, 0, 80, 90, 30).Name = "Scroll"
ActiveSheet.Shapes.AddFormControl(4, 0, 120, 90, 30).Name = "Group"
End Sub''', name="Builder")
    app.run("Build")
    out = tmp_path / "issue25.xlsm"
    app.save(out)
    with harness.ExcelSession(harness.HarnessConfig(lock_wait_s=45.0)) as excel:
        excel.open_document(out)
        result = excel.run_vba('''Public Function Report() As String
Dim s As Object
Set s = ActiveSheet.Shapes
s("Spin").ControlFormat.Value = 7
s("Scroll").ControlFormat.Value = 25
Report = CStr(s("Radio").FormControlType) & "|" & CStr(s("Radio").ControlFormat.Value) & "|" & _
CStr(s("Spin").FormControlType) & "|" & CStr(s("Spin").ControlFormat.Max) & "|" & CStr(s("Spin").ControlFormat.Value) & "|" & _
CStr(s("Scroll").FormControlType) & "|" & CStr(s("Scroll").ControlFormat.Max) & "|" & CStr(s("Scroll").ControlFormat.Value) & "|" & _
CStr(s("Group").FormControlType)
End Function''', "Report", timeout=120.0)
        assert result.ok, f"{result.outcome}: {result.message}"
        assert str(result.value) == "7|-4146|9|30000|7|8|100|25|4"


def test_excel_reads_numeric_control_edits(tmp_path: Path) -> None:
    harness = pytest.importorskip("pyvbaharness")
    app = ExcelApplication.open(Path(__file__).parent / "fixtures/shapes/all_controls.xlsm", with_vba=False)
    app.add_module('''Public Sub Edit()
Dim sh As Object
For Each sh In ActiveSheet.Shapes
If sh.Name = "Step" Or sh.Name = "Slide" Then
sh.ControlFormat.Min = 5
sh.ControlFormat.Max = 45
sh.ControlFormat.SmallChange = 2
sh.ControlFormat.Value = 13
End If
Next sh
End Sub''', name="Editor")
    app.run("Edit")
    out = tmp_path / "numeric_edits.xlsm"
    app.save(out)
    with harness.ExcelSession(harness.HarnessConfig(lock_wait_s=45.0)) as excel:
        excel.open_document(out)
        result = excel.run_vba('''Public Function Report() As String
Dim sh As Object, cf As Object
For Each sh In ActiveSheet.Shapes
If sh.Name = "Step" Or sh.Name = "Slide" Then
Set cf = sh.ControlFormat
Report = Report & CStr(cf.Value) & "|" & CStr(cf.Min) & "|" & CStr(cf.Max) & "|" & CStr(cf.SmallChange) & ";"
End If
Next sh
End Function''', "Report", timeout=120.0)
        assert result.ok, f"{result.outcome}: {result.message}"
        assert str(result.value) == "13|5|45|2;13|5|45|2;"


def test_excel_reads_public_control_creation_and_rebinding(tmp_path: Path) -> None:
    harness = pytest.importorskip("pyvbaharness")
    app = ExcelApplication()
    app.add_workbook()
    sheet = app.sheet(1)
    kinds = [0, 1, 2, 4, 5, 6, 7, 8, 9]
    for kind in kinds:
        sheet.add_form_control(kind, name=f"Control{kind}", top=kind * 40, text=f"Caption {kind}")
    sheet.set_control_items("Control6", ["a", "b", "c"])
    sheet.set_control_selection_mode("Control6", 2)
    sheet.set_control_selection("Control6", [1, 3])
    sheet.set_control_selection_mode("Control6", 3)
    sheet.set_value("J1", "x")
    sheet.set_value("J2", "y")
    sheet.update_control("Control6", list_range="$J$1:$J$2")
    path = tmp_path / "public_controls.xlsm"
    app.save(path)
    with harness.ExcelSession(harness.HarnessConfig(lock_wait_s=45.0)) as excel:
        excel.open_document(path)
        result = excel.run_vba('''Public Function Report() As String
Dim sh As Object
For Each sh In ActiveSheet.Shapes
Report = Report & sh.Name & ":" & CStr(sh.FormControlType) & ";"
Next sh
Set sh = ActiveSheet.Shapes("Control6")
Report = Report & "|" & CStr(sh.ControlFormat.MultiSelect) & "|" & sh.ControlFormat.ListFillRange & "|" & _
CStr(sh.ControlFormat.ListCount) & "|" & CStr(sh.DrawingObject.Selected(1)) & "|" & CStr(sh.DrawingObject.Selected(2))
End Function''', "Report", timeout=120.0)
        assert result.ok, f"{result.outcome}: {result.message}"
        expected = "".join(f"Control{kind}:{kind};" for kind in kinds) + "|3|$J$1:$J$2|2|True|False"
        assert str(result.value) == expected


@pytest.mark.parametrize("layout,reverse", [("overlap", False), ("overlap", True),
                                           ("nested", False), ("nested", True)])
def test_excel_reads_overlapping_radio_groups(tmp_path: Path, layout: str, reverse: bool) -> None:
    import json

    harness = pytest.importorskip("pyvbaharness")
    records = json.loads((Path(__file__).parent / "fixtures/shapes/radio_overlap.json").read_text())
    record = next(r for r in records if r["layout"] == layout and r["reverse"] == reverse and r["operation"] == "select")
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('Public Function Report() As String\nDim sh As Object, box1 As Object, box2 As Object\n'
                   + record["body"] + 'End Function\n', name="Builder")
    assert app.run("Report") == record["reported"]
    path = tmp_path / "overlapping.xlsm"
    app.save(path)
    app = ExcelApplication.open(path, with_vba=False)
    inspect = '''Public Function Inspect() As String
Dim sh As Object
Inspect = CStr(Range("H1").Value) & ":" & CStr(Range("H2").Value) & ":" & CStr(Range("H3").Value) & "|"
For Each sh In ActiveSheet.Shapes
If Left(sh.Name, 5) = "Radio" Then
Inspect = Inspect & sh.Name & ":" & CStr(sh.ControlFormat.Value) & ":" & sh.ControlFormat.LinkedCell & ";"
End If
Next sh
End Function'''
    app.add_module(inspect, name="Inspector")
    assert app.run("Inspect") == record["reported"]
    app.sheet(1).set_control_value("Radio2", 0)
    app.sheet(1).set_control_value("Radio2", 1)
    expected = app.run("Inspect")
    app.save(path)
    with harness.ExcelSession(harness.HarnessConfig(lock_wait_s=45.0)) as excel:
        excel.open_document(path)
        result = excel.run_vba(inspect, "Inspect", timeout=120.0)
        assert result.ok, f"{result.outcome}: {result.message}"
        assert str(result.value) == expected


def test_excel_reads_headless_radio_groups(tmp_path: Path) -> None:
    harness = pytest.importorskip("pyvbaharness")
    app = ExcelApplication()
    app.add_workbook()
    sheet = app.sheet(1)
    sheet.add_form_control(4, width=150, height=100)
    sheet.add_form_control(7, name="First", left=10, top=10, width=90, height=20)
    sheet.add_form_control(7, name="Second", left=10, top=40, width=90, height=20)
    sheet.add_form_control(7, name="Outside", left=200, top=10, width=90, height=20)
    sheet.update_control("Second", linked_cell="$H$1")
    sheet.set_control_value("Second", 1)
    sheet.set_control_value("Outside", 1)
    path = tmp_path / "radio_groups.xlsm"
    app.save(path)
    # Exercise both newly generated and existing supporting control parts.
    app = ExcelApplication.open(path, with_vba=False)
    app.sheet(1).set_control_value("First", 1)
    app.save(path)
    with harness.ExcelSession(harness.HarnessConfig(lock_wait_s=45.0)) as excel:
        excel.open_document(path)
        result = excel.run_vba('''Public Function Report() As String
Dim a As Object, b As Object, c As Object
Set a = ActiveSheet.Shapes("First").ControlFormat
Set b = ActiveSheet.Shapes("Second").ControlFormat
Set c = ActiveSheet.Shapes("Outside").ControlFormat
Report = CStr(a.Value) & "|" & CStr(b.Value) & "|" & CStr(c.Value) & "|" & CStr(Range("H1").Value) & "|" & b.LinkedCell
b.Value = 1
Report = Report & "|" & CStr(a.Value) & "|" & CStr(b.Value) & "|" & CStr(c.Value) & "|" & CStr(Range("H1").Value)
End Function''', "Report", timeout=120.0)
        assert result.ok, f"{result.outcome}: {result.message}"
        assert str(result.value) == "1|-4146|1|1|$H$1|-4146|1|1|2"


@pytest.mark.parametrize("operation", ["add", "delete"])
def test_excel_reads_regrouped_radios(tmp_path: Path, operation: str) -> None:
    import json

    harness = pytest.importorskip("pyvbaharness")
    records = json.loads((Path(__file__).parent / "fixtures/shapes/radio_regrouping.json").read_text())
    record = next(r for r in records if r["operation"] == operation and r["chosen"] == 2 and r["outside"])
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('Public Function Report() As String\nDim a As Object, b As Object, c As Object, box As Object\n'
                   + record["body"] + 'End Function\n', name="Builder")
    assert app.run("Report") == record["reported"]
    path = tmp_path / "regrouped.xlsm"
    app.save(path)
    with harness.ExcelSession(harness.HarnessConfig(lock_wait_s=45.0)) as excel:
        excel.open_document(path)
        result = excel.run_vba('''Public Function Inspect() As String
Dim a As Object, b As Object, c As Object
Set a = ActiveSheet.Shapes("First").ControlFormat
Set b = ActiveSheet.Shapes("Second").ControlFormat
Set c = ActiveSheet.Shapes("Third").ControlFormat
Inspect = CStr(a.Value) & ":" & a.LinkedCell & ";" & CStr(b.Value) & ":" & b.LinkedCell & ";" & _
CStr(c.Value) & ":" & c.LinkedCell & ";" & CStr(Range("H1").Value) & ":" & CStr(Range("H2").Value)
End Function''', "Inspect", timeout=120.0)
        assert result.ok, f"{result.outcome}: {result.message}"
        assert str(result.value) == record["reported"].split("|")[-1]


@pytest.mark.parametrize("local", [False, True])
def test_excel_reads_named_control_bindings(tmp_path: Path, local: bool) -> None:
    harness = pytest.importorskip("pyvbaharness")
    app = ExcelApplication()
    app.add_workbook()
    local_name = 'ActiveWorkbook.Names.Add "Sheet1!Target", "=Sheet1!$H$2"\n' if local else ""
    app.add_module('Public Sub Build()\n'
                   'ActiveWorkbook.Names.Add "Target", "=Sheet1!$H$1"\n'
                   'ActiveWorkbook.Names.Add "Choices", "=Sheet1!$J$1:$J$3"\n'
                   + local_name + 'End Sub\n', name="Builder")
    app.run("Build")
    sheet = app.sheet(1)
    for row, value in enumerate(["a", "b", "c"], 1):
        sheet.set_value(f"J{row}", value)
    sheet.add_form_control(6, name="Picker")
    sheet.update_control("Picker", linked_cell="=target", list_range="choices")
    sheet.set_control_value("Picker", 2)
    path = tmp_path / "named_controls.xlsm"
    app.save(path)
    app = ExcelApplication.open(path, with_vba=False)
    app.sheet(1).set_control_value("Picker", 3)
    app.save(path)
    with harness.ExcelSession(harness.HarnessConfig(lock_wait_s=45.0)) as excel:
        excel.open_document(path)
        result = excel.run_vba('''Public Function Inspect() As String
Dim cf As Object
Set cf = ActiveSheet.Shapes("Picker").ControlFormat
Inspect = cf.LinkedCell & "|" & cf.ListFillRange & "|" & CStr(cf.Value) & "|" & CStr(cf.ListCount) & "|" & _
CStr(Range("H1").Value) & "|" & CStr(Range("H2").Value) & "|" & cf.List(2)
End Function''', "Inspect", timeout=120.0)
        assert result.ok, f"{result.outcome}: {result.message}"
        assert str(result.value) == ("Target|Choices|3|3||3|b" if local else "Target|Choices|3|3|3||b")


@pytest.mark.parametrize("layout", ["identical", "nested", "outer", "partial"])
def test_excel_reads_late_overlapping_boxes(tmp_path: Path, layout: str) -> None:
    import json

    harness = pytest.importorskip("pyvbaharness")
    records = json.loads((Path(__file__).parent / "fixtures/shapes/radio_late_boxes.json").read_text())
    record = next(r for r in records if r["layout"] == layout and r["links"] == "inside" and r["chosen"] == 1)
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('Public Function Report() As String\nDim a As Object, b As Object, c As Object\n'
                   + record["body"] + 'End Function\n', name="Builder")
    assert app.run("Report") == record["reported"]
    path = tmp_path / "late_boxes.xlsm"
    app.save(path)
    inspect = '''Public Function Inspect() As String
Dim a As Object, b As Object, c As Object
Set a = ActiveSheet.Shapes("First").ControlFormat
Set b = ActiveSheet.Shapes("Second").ControlFormat
Set c = ActiveSheet.Shapes("Third").ControlFormat
Inspect = CStr(a.Value) & ":" & a.LinkedCell & ";" & CStr(b.Value) & ":" & b.LinkedCell & ";" & _
CStr(c.Value) & ":" & c.LinkedCell & ";" & CStr(Range("H1").Value) & ":" & CStr(Range("H2").Value)
End Function'''
    reopened = ExcelApplication.open(path, with_vba=False)
    reopened.add_module(inspect, name="Inspector")
    assert reopened.run("Inspect") == record["reported"].split("|")[-1]
    with harness.ExcelSession(harness.HarnessConfig(lock_wait_s=45.0)) as excel:
        excel.open_document(path)
        result = excel.run_vba(inspect, "Inspect", timeout=120.0)
        assert result.ok, f"{result.outcome}: {result.message}"
        assert str(result.value) == record["reported"].split("|")[-1]


@pytest.mark.parametrize("operation", ["select", "delete_first_box"])
def test_excel_reads_interleaved_radio_groups(tmp_path: Path, operation: str) -> None:
    import json

    harness = pytest.importorskip("pyvbaharness")
    records = json.loads((Path(__file__).parent / "fixtures/shapes/radio_interleaving.json").read_text())
    record = next(r for r in records if r["boxes"] == 2 and r["order"] == "ABCABC" and r["operation"] == operation)
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('Public Function Report() As String\nDim sh As Object, box1 As Object, box2 As Object\n'
                   + record["body"] + 'End Function\n', name="Builder")
    assert app.run("Report") == record["reported"]
    path = tmp_path / "interleaved.xlsm"
    app.save(path)
    app = ExcelApplication.open(path, with_vba=False)
    inspect = '''Public Function Inspect() As String
Dim sh As Object
Inspect = CStr(Range("H1").Value) & ":" & CStr(Range("H2").Value) & "|"
For Each sh In ActiveSheet.Shapes
If Left(sh.Name, 5) = "Radio" Then
Inspect = Inspect & sh.Name & ":" & CStr(sh.ControlFormat.Value) & ":" & sh.ControlFormat.LinkedCell & ";"
End If
Next sh
End Function'''
    app.add_module(inspect, name="Inspector")
    assert app.run("Inspect") == record["reported"]
    app.sheet(1).set_control_value("Radio4", 1)
    expected = app.run("Inspect")
    app.save(path)
    with harness.ExcelSession(harness.HarnessConfig(lock_wait_s=45.0)) as excel:
        excel.open_document(path)
        result = excel.run_vba(inspect, "Inspect", timeout=120.0)
        assert result.ok, f"{result.outcome}: {result.message}"
        assert str(result.value) == expected


@pytest.mark.parametrize("formula", ["OFFSET(Sheet1!$J$1,0,0,Sheet1!$K$1,1)",
                                    'INDIRECT(Sheet1!$K$2)'])
def test_excel_reads_dynamic_named_controls(tmp_path: Path, formula: str) -> None:
    harness = pytest.importorskip("pyvbaharness")
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(f'''Public Sub Build()
ActiveWorkbook.Names.Add "Dynamic", "={formula}"
ActiveWorkbook.Names.Add "Target", "=OFFSET(Sheet1!$H$1,1,0)"
End Sub''', name="Builder")
    app.run("Build")
    sheet = app.sheet(1)
    sheet.set_value("J1", "a")
    sheet.set_value("J2", "b")
    sheet.set_value("J3", "c")
    sheet.set_value("K1", 3)
    sheet.set_value("K2", "Sheet1!$J$1:$J$3")
    sheet.add_form_control(6, name="Picker")
    sheet.update_control("Picker", list_range="Dynamic", linked_cell="Target")
    sheet.set_control_value("Picker", 3)
    path = tmp_path / "dynamic_names.xlsm"
    app.save(path)
    app = ExcelApplication.open(path, with_vba=False)
    sheet = app.sheet(1)
    sheet.set_value("K1", 2)
    sheet.set_value("K2", "Sheet1!$J$1:$J$2")
    sheet.set_control_value("Picker", 1)
    sheet.set_control_value("Picker", 2)
    assert sheet.control_items("Picker") == ["a", "b"]
    assert sheet.value("H2") == 2
    app.save(path)
    inspect = '''Public Function Inspect() As String
Dim cf As Object
Set cf = ActiveSheet.Shapes("Picker").ControlFormat
Inspect = cf.ListFillRange & "|" & cf.LinkedCell & "|" & CStr(cf.ListCount) & "|" & CStr(cf.Value) & "|" & CStr(Range("H2").Value)
Range("K1").Value = 3
Range("K2").Value = "Sheet1!$J$1:$J$3"
cf.Value = 3
Inspect = Inspect & "|" & CStr(cf.ListCount) & "|" & cf.List(3) & "|" & CStr(Range("H2").Value)
End Function'''
    with harness.ExcelSession(harness.HarnessConfig(lock_wait_s=45.0)) as excel:
        excel.open_document(path)
        result = excel.run_vba(inspect, "Inspect", timeout=120.0)
        assert result.ok, f"{result.outcome}: {result.message}"
        assert str(result.value) == "Dynamic|Target|2|2|2|3|c|3"


@pytest.mark.parametrize("reference", ["index", "choose", "if", "indirect_r1c1"])
def test_excel_reads_branching_named_controls(tmp_path: Path, reference: str) -> None:
    harness = pytest.importorskip("pyvbaharness")
    app = ExcelApplication()
    app.add_workbook()
    sources = {
        "index": "INDEX(Sheet1!$J$1:$K$3,0,Sheet1!$M$1)",
        "choose": "CHOOSE(Sheet1!$M$1,Sheet1!$J$1:$J$3,Sheet1!$K$1:$K$3)",
        "if": "IF(Sheet1!$M$1=1,Sheet1!$J$1:$J$3,Sheet1!$K$1:$K$3)",
        "indirect_r1c1": 'INDIRECT(""Sheet1!R1C""&(9+Sheet1!$M$1)&"":R3C""&(9+Sheet1!$M$1),FALSE)',
    }
    targets = {
        "index": "INDEX(Sheet1!$H$1:$H$2,Sheet1!$M$2)",
        "choose": "CHOOSE(Sheet1!$M$2,Sheet1!$H$1,Sheet1!$H$2)",
        "if": "IF(Sheet1!$M$2=1,Sheet1!$H$1,Sheet1!$H$2)",
        "indirect_r1c1": 'INDIRECT(""Sheet1!R""&Sheet1!$M$2&""C8"",FALSE)',
    }
    app.add_module(f'''Public Sub Build()
ActiveWorkbook.Names.Add "Choices", "={sources[reference]}"
ActiveWorkbook.Names.Add "Target", "={targets[reference]}"
End Sub''', name="Builder")
    app.run("Build")
    sheet = app.sheet(1)
    for address, value in {"J1": "a", "J2": "b", "J3": "c", "K1": "x", "K2": "y", "K3": "z",
                           "M1": 1, "M2": 1}.items():
        sheet.set_value(address, value)
    sheet.add_form_control(6, name="Picker")
    sheet.update_control("Picker", list_range="Choices", linked_cell="Target")
    sheet.set_control_value("Picker", 3)
    path = tmp_path / "branching_names.xlsm"
    app.save(path)
    app = ExcelApplication.open(path, with_vba=False)
    sheet = app.sheet(1)
    sheet.set_value("M1", 2)
    sheet.set_value("M2", 2)
    sheet.set_control_value("Picker", 2)
    assert sheet.control_items("Picker") == ["x", "y", "z"]
    assert sheet.value("H1") == 3
    assert sheet.value("H2") == 2
    app.save(path)
    inspect = '''Public Function Inspect() As String
Dim cf As Object
Set cf = ActiveSheet.Shapes("Picker").ControlFormat
Inspect = cf.ListFillRange & "|" & cf.LinkedCell & "|" & cf.List(2) & "|" & CStr(cf.Value) & "|" & CStr(Range("H2").Value)
Range("M1").Value = 1
Range("M2").Value = 1
cf.Value = 1
Inspect = Inspect & "|" & cf.List(2) & "|" & CStr(Range("H1").Value)
End Function'''
    with harness.ExcelSession(harness.HarnessConfig(lock_wait_s=45.0)) as excel:
        excel.open_document(path)
        result = excel.run_vba(inspect, "Inspect", timeout=120.0)
        assert result.ok, f"{result.outcome}: {result.message}"
        assert str(result.value) == "Choices|Target|y|2|2|b|1"


def test_excel_find_on_headless_saved_cells(tmp_path: Path) -> None:
    harness = pytest.importorskip("pyvbaharness")
    app = ExcelApplication()
    app.add_workbook()
    setup = '''Public Sub Build()
Range("A1").Value = "Alpha"
Range("B1").Value = "alphabet"
Range("A2").Value = "alpha"
Range("C3").Formula = "=1/2"
Range("C3").NumberFormat = "0%"
End Sub'''
    report = '''Public Function Report() As String
Dim hit As Object
Set hit = Range("A1:C3").Find("alpha", , xlFormulas, xlPart, xlByRows, xlNext, False, False, False)
Report = hit.Address
Set hit = Range("A1:C3").FindNext(hit)
Report = Report & "|" & hit.Address
Set hit = Range("A1:C3").Find("50%", LookIn:=xlValues)
Report = Report & "|" & hit.Address
End Function'''
    app.add_module(setup, name="Builder")
    app.add_module(report, name="Probe")
    app.run("Build")
    expected = app.run("Report")
    assert expected == "$B$1|$A$2|$C$3"
    path = tmp_path / "find.xlsm"
    app.save(path)
    reopened = ExcelApplication.open(path)
    reopened.add_module(report, name="Probe")
    assert reopened.run("Report") == expected
    with harness.ExcelSession(harness.HarnessConfig(lock_wait_s=45.0)) as excel:
        excel.open_document(path)
        result = excel.run_vba(report, "Report", timeout=120.0)
        assert result.ok, result.message
        assert result.value == expected


@pytest.mark.parametrize("across", [False, True])
@pytest.mark.parametrize("unmerge", [False, True])
def test_excel_reads_merged_ranges(tmp_path: Path, across: bool, unmerge: bool) -> None:
    harness = pytest.importorskip("pyvbaharness")
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(f'''Public Sub Build()
Range("C2").Value = "first"
Range("B3").Value = "second"
Range("B2:D4").Merge Across:={str(across)}
If {str(unmerge)} Then Range("C3").UnMerge
End Sub''', name="Builder")
    report = '''Public Function Report() As String
Report = Range("B2").MergeArea.Address & "|" & Range("C3").MergeArea.Address & "|" & CStr(Range("B2").Value) & "|" & CStr(Range("B3").Value)
End Function'''
    app.add_module(report, name="Probe")
    app.run("Build")
    expected = app.run("Report")
    path = tmp_path / "merged_ranges.xlsm"
    app.save(path)
    reopened = ExcelApplication.open(path, with_vba=False)
    reopened.add_module(report, name="Probe")
    assert reopened.run("Report") == expected
    with harness.ExcelSession(harness.HarnessConfig(lock_wait_s=45.0)) as excel:
        excel.open_document(path)
        result = excel.run_vba(report, "Report", timeout=120.0)
        assert result.ok, result.message
        assert result.value == expected


def test_excel_reads_r1c1_formulas(tmp_path: Path) -> None:
    harness = pytest.importorskip("pyvbaharness")
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('''Sub Build()
Range("A1:A3").Value = 7
Range("B1:B3").FormulaR1C1 = "=RC[-1]*2+R1C1"
Application.Calculate
End Sub''', name="Builder")
    report = '''Function Report() As String
Report = Range("B3").Formula & "|" & Range("B3").FormulaR1C1 & "|" & CStr(Range("B3").Value)
End Function'''
    app.add_module(report, name="Probe")
    app.run("Build")
    expected = app.run("Report")
    assert expected == "=A3*2+$A$1|=RC[-1]*2+R1C1|21"
    path = tmp_path / "r1c1.xlsm"
    app.save(path)
    with harness.ExcelSession(harness.HarnessConfig(lock_wait_s=45.0)) as excel:
        excel.open_document(path)
        result = excel.run_vba(report, "Report", timeout=120.0)
        assert result.ok, result.message
        assert result.value == expected


@pytest.mark.parametrize("name", ["row_matrix", "tiled", "errors", "lower_bounds"])
@pytest.mark.parametrize("style", ["a1", "r1c1"])
def test_excel_reads_formula_arrays(tmp_path: Path, name: str, style: str) -> None:
    import json

    harness = pytest.importorskip("pyvbaharness")
    filename = "formula_a1_arrays.json" if style == "a1" else "formula_arrays.json"
    records = json.loads((Path(__file__).parent / "fixtures" / filename).read_text())
    record = next(r for r in records if r["name"] == name)
    app = ExcelApplication()
    app.add_workbook()
    body = record["body"]
    app.add_module('Public Function Build() As String\nDim cell As Object, n As Long\n'
                   + body.replace("Report", "Build") + 'End Function', name="Builder")
    assert app.run("Build") == record["reported"]
    path = tmp_path / "formula_arrays.xlsm"
    app.save(path)
    report = ('Public Function Report() As String\nDim cell As Object, n As Long\n'
              + body[body.index('Report = CStr(n)'):] + 'End Function')
    with harness.ExcelSession(harness.HarnessConfig(lock_wait_s=45.0)) as excel:
        excel.open_document(path)
        result = excel.run_vba(report, "Report", timeout=120.0)
        assert result.ok, result.message
        assert result.value == record["reported"]


@pytest.mark.parametrize("name", ["overlap_down", "tile", "blank_tail", "tiled_formula", "format"])
def test_excel_reads_copied_ranges(tmp_path: Path, name: str) -> None:
    import json

    harness = pytest.importorskip("pyvbaharness")
    records = json.loads((Path(__file__).parent / "fixtures/range_copy.json").read_text())
    record = next(r for r in records if r["name"] == name)
    app = ExcelApplication()
    app.add_workbook()
    body = record["body"]
    app.add_module('Function Build() As String\nDim cell As Object, n As Long\n'
                   + body.replace("Report", "Build") + 'End Function', name="Builder")
    assert app.run("Build") == record["reported"]
    path = tmp_path / "copied_ranges.xlsm"
    app.save(path)
    report = ('Function Report() As String\nDim cell As Object, n As Long\n'
              + body[body.index('Report = CStr(n)'):] + 'End Function')
    with harness.ExcelSession(harness.HarnessConfig(lock_wait_s=45.0)) as excel:
        excel.open_document(path)
        result = excel.run_vba(report, "Report", timeout=120.0)
        assert result.ok, result.message
        assert result.value == record["reported"]


@pytest.mark.parametrize("name", ["insert_row", "delete_row", "insert_column", "delete_columns"])
def test_excel_reads_structural_edits(tmp_path: Path, name: str) -> None:
    import json

    harness = pytest.importorskip("pyvbaharness")
    records = json.loads((Path(__file__).parent / "fixtures/range_edit.json").read_text())
    record = next(r for r in records if r["name"] == name)
    app = ExcelApplication()
    app.add_workbook()
    body = record["body"]
    app.add_module('Function Build() As String\nDim cell As Object, n As Long\n'
                   + body.replace("Report", "Build") + 'End Function', name="Builder")
    assert app.run("Build") == record["reported"]
    path = tmp_path / "structural_edits.xlsm"
    app.save(path)
    report = ('Function Report() As String\nDim cell As Object, n As Long\n'
              + 'Worksheets("Sheet1").Activate\n' + body[body.index('Report = CStr(n)'):] + 'End Function')
    with harness.ExcelSession(harness.HarnessConfig(lock_wait_s=45.0)) as excel:
        excel.open_document(path)
        result = excel.run_vba(report, "Report", timeout=120.0)
        assert result.ok, result.message
        assert result.value == record["reported"]


@pytest.mark.parametrize("stage", ["created", "updated", "deleted"])
def test_excel_named_range_crud_persistence(tmp_path: Path, stage: str) -> None:
    harness = pytest.importorskip("pyvbaharness")
    app = ExcelApplication()
    app.add_workbook()
    app.sheet(1).set_value("A1", 10)
    app.sheet(1).set_value("B1", 20)
    app.add_named_range("GlobalName", "=Sheet1!$A$1", visible=False, comment="global")
    app.sheet(1).add_named_range("LocalName", "=$B$1", comment="local")
    app.sheet(1).set_value("C1", "=GlobalName+LocalName")
    path = tmp_path / "named_ranges.xlsm"
    app.save(path)
    app = ExcelApplication.open(path, with_vba=False)
    if stage != "created":
        app.update_named_range("GlobalName", new_name="RenamedGlobal", refers_to="=Sheet1!$B$1", visible=True, comment="changed")
        app.sheet(1).update_named_range("LocalName", new_name="RenamedLocal", refers_to="=$A$1", visible=False)
    if stage == "deleted":
        app.remove_named_range("RenamedGlobal")
        app.sheet(1).remove_named_range("RenamedLocal")
    global_name, local_name = ("GlobalName", "LocalName") if stage == "created" else ("RenamedGlobal", "RenamedLocal")
    # Excel may create an internal _xlfn.SINGLE compatibility name when it
    # opens a formula containing a deleted name. Count the user definitions.
    report = '''Function Report() As String
Dim nm As Object, userCount As Long
For Each nm In ThisWorkbook.Names
If Left(nm.Name, 6) <> "_xlfn." Then userCount = userCount + 1
Next nm
Report = CStr(userCount) & "|" & CStr(Worksheets(1).Names.Count) & "|" & Range("C1").Formula & "|" & CStr(Range("C1").Value)
'''
    if stage != "deleted":
        report += f'''Report = Report & "|" & ThisWorkbook.Names("{global_name}").RefersTo & "|" & CStr(ThisWorkbook.Names("{global_name}").Visible) & "|" & ThisWorkbook.Names("{global_name}").Comment
Report = Report & "|" & Worksheets(1).Names("{local_name}").Name & "|" & Worksheets(1).Names("{local_name}").RefersTo
'''
    report += "End Function"
    app.add_module(report, name="Probe")
    expected = app.run("Report")
    app.save(path)
    with harness.ExcelSession(harness.HarnessConfig(lock_wait_s=45.0)) as excel:
        excel.open_document(path)
        result = excel.run_vba(report, "Report", timeout=120.0)
        assert result.ok, result.message
        assert result.value == expected


@pytest.mark.parametrize("reopen", [False, True])
def test_excel_cross_workbook_copy_persistence(tmp_path: Path, reopen: bool) -> None:
    harness = pytest.importorskip("pyvbaharness")
    app = ExcelApplication()
    app.add_workbook()
    source = app.sheet(1)
    source.set_value("A1", 10)
    source.set_value("B1", "=A1+1")
    source.add_named_range("LocalAmount", "=$A$1")
    destination = app.add_workbook()
    path = tmp_path / "copied_book.xlsx"
    if reopen:
        app.save(path, workbook=destination)
        destination.Close(SaveChanges=False)
        destination = app.open_workbook(path)
    source.copy(after=app.sheet(1, workbook=destination))
    source.copy_range("A1:B1", app.sheet(1, workbook=destination), "C3")
    app.save(path, workbook=destination)
    report = '''Function Report() As String
Report = CStr(ThisWorkbook.Worksheets.Count) & "|" & Worksheets(2).Name & "|" & CStr(Worksheets(2).Range("B1").Value) & "|" & Worksheets(1).Range("D3").Formula & "|" & CStr(Worksheets(1).Range("D3").Value) & "|" & Worksheets(2).Names("LocalAmount").RefersTo
End Function'''
    with harness.ExcelSession(harness.HarnessConfig(lock_wait_s=45.0)) as excel:
        excel.open_document(path)
        result = excel.run_vba(report, "Report", timeout=120)
        assert result.ok, result.message
        assert result.value == "2|Sheet1 (2)|11|=C3+1|11|='Sheet1 (2)'!$A$1"


@pytest.mark.parametrize("local", [False, True])
def test_excel_range_copy_name_deconfliction(tmp_path: Path, local: bool) -> None:
    harness = pytest.importorskip("pyvbaharness")
    app = ExcelApplication()
    app.add_workbook()
    source = app.sheet(1)
    names = source if local else app
    names.add_named_range("Amount", "=Sheet1!$A$1", visible=False, comment="imported")
    names.add_named_range("Total", "=Amount")
    source.set_value("B1", "=Total+Amount")
    app.add_workbook()
    destination = app.sheet(1)
    destination.set_value("A1", 7)
    app.add_named_range("Amount", "=99")
    source.copy_range("B1", destination, "D3", name_conflict="rename")
    path = tmp_path / "deconflicted.xlsx"
    app.save(path)
    report = '''Function Report() As String
Report = Range("D3").Formula & "|" & CStr(Range("D3").Value) & "|" & ThisWorkbook.Names("Amount").RefersTo
End Function'''
    with harness.ExcelSession(harness.HarnessConfig(lock_wait_s=45.0)) as excel:
        excel.open_document(path)
        result = excel.run_vba(report, "Report", timeout=120)
        assert result.ok, result.message
        assert result.value == "=Total+Amount_2|14|=99"


@pytest.mark.parametrize("cross_book", [False, True])
def test_excel_worksheet_move_persistence(tmp_path: Path, cross_book: bool) -> None:
    harness = pytest.importorskip("pyvbaharness")
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('''Sub Build()
Dim sh As Object
Set sh = ThisWorkbook.Worksheets(1)
sh.Range("A1").Value = 7
sh.Names.Add "LocalAmount", "=Sheet1!$A$1"
sh.Range("B1").Formula = "=LocalAmount+1"
ThisWorkbook.Worksheets.Add(After:=sh).Name = "Other"
End Sub''', name="Builder")
    app.run("Build")
    path = tmp_path / "moved.xlsx"
    app.save(path)
    statement = ('Set dst = Workbooks.Add\nThisWorkbook.Worksheets(1).Move After:=dst.Worksheets(1)'
                 if cross_book else 'ThisWorkbook.Worksheets(1).Move After:=ThisWorkbook.Worksheets(2)')
    app.add_module('Sub Relocate()\nDim dst As Object\n' + statement + '\nEnd Sub', name="Mover")
    app.run("Relocate")
    app.save(path)
    report = '''Function Report() As String
Report = Worksheets(1).Name & "|" & Worksheets(2).Name & "|" & CStr(Worksheets(2).Range("B1").Value) & "|" & Worksheets(2).Names("LocalAmount").RefersTo
End Function'''
    with harness.ExcelSession(harness.HarnessConfig(lock_wait_s=45.0)) as excel:
        excel.open_document(path)
        result = excel.run_vba(report, "Report", timeout=120)
        assert result.ok, result.message
        expected = "Sheet1|Sheet1 (2)|8|='Sheet1 (2)'!$A$1" if cross_book else "Other|Sheet1|8|=Sheet1!$A$1"
        assert result.value == expected


def test_excel_reads_model_authored_formats(tmp_path: Path) -> None:
    """Every formatting case the model writes reads back in Excel as Excel's own file of it does."""
    import json

    harness = pytest.importorskip("pyvbaharness")
    fixture = json.loads((Path(__file__).parent / "fixtures" / "format" / "formats_answers.json").read_text())
    cases: list[str] = fixture["cases"]
    app = ExcelApplication()
    app.add_workbook()
    lines = ["Public Sub Build()", "Dim c As Object"]
    for row, name in enumerate(cases, start=1):
        lines += [f'Set c = ActiveSheet.Range("A{row}")', f'ActiveSheet.Range("B{row}").Value = "{name}"']
        lines += [line.strip() for line in str(fixture["setups"][name]).splitlines()]
    app.add_module("\n".join([*lines, "End Sub"]) + "\n", name="Builder")
    app.run("Build")
    path = tmp_path / "model_formats.xlsx"
    app.save(path)
    reader = (fixture["describe"] + "Public Function Report() As String\nDim row As Long\n"
              f"For row = 1 To {len(cases)}\n"
              'Report = Report & Describe(ActiveSheet.Cells(row, 1)) & "|"\nNext row\nEnd Function\n')
    with harness.ExcelSession(harness.HarnessConfig(lock_wait_s=45.0)) as excel:
        excel.open_document(path)
        result = excel.run_vba(reader, "Report", timeout=240.0)
        assert result.ok, result.message
    reads: list[str] = fixture["reads"]
    for name, described in zip(cases, str(result.value).split("|")[:-1], strict=True):
        assert dict(zip(reads, described.split(";")[:-1], strict=True)) == fixture["answers"][name], name


def test_excel_reads_model_authored_sizes(tmp_path: Path) -> None:
    """Every row and column sizing case the model saves reads in Excel as Excel's own file of it does.

    The model sizes rows and columns for a 96-DPI display; on any other
    Excel rounds to other pixels, so the gate says so rather than fail.
    """
    import json

    harness = pytest.importorskip("pyvbaharness")
    fixture = json.loads((Path(__file__).parent / "fixtures" / "dimensions" / "dimensions_answers.json").read_text())
    cases: list[dict[str, str]] = fixture["cases"]
    app = ExcelApplication()
    app.add_workbook()
    lines = ["Public Sub Build()", "Dim ws As Object"]
    for index, case in enumerate(cases):
        lines.append("Set ws = ActiveWorkbook.Worksheets(1)" if index == 0 else
                     "Set ws = ActiveWorkbook.Worksheets.Add(After:=ActiveWorkbook.Worksheets"
                     "(ActiveWorkbook.Worksheets.Count))")
        lines += [f'ws.Name = "{case["name"]}"', *case["setup"].splitlines()]
    app.add_module("\n".join([*lines, "ActiveWorkbook.Worksheets(1).Activate", "End Sub"]) + "\n", name="Builder")
    app.run("Build")
    path = tmp_path / "model_sizes.xlsx"
    app.save(path)
    reader = (fixture["describe"] + "Public Function Report() As String\nDim ws As Object\n"
              'Report = ActiveSheet.StandardHeight & "#"\n'
              'For Each ws In ActiveWorkbook.Worksheets\nReport = Report & Describe(ws) & "|"\nNext ws\nEnd Function\n')
    with harness.ExcelSession(harness.HarnessConfig(lock_wait_s=45.0)) as excel:
        excel.open_document(path)
        result = excel.run_vba(reader, "Report", timeout=240.0)
        assert result.ok, result.message
    standard, described = str(result.value).split("#", 1)
    if standard != "15":
        pytest.skip(f"Excel is on a display whose standard row is {standard}pt; the model emulates 96 DPI (15pt)")
    for case, answer in zip(cases, described.split("|")[:-1], strict=True):
        # A recorded_ case rests on row heights only Excel works out; the
        # model saves those rows without one, and Excel does not always
        # work it out again when it opens the file.
        if not case["name"].startswith("recorded_"):
            assert answer == case["answers"], case["name"]


def _row_format_reads(name: str, reads: list[str]) -> str:
    """The reads scripts/measure_row_formats.py made of one case: the used range first, then each read."""
    lines = [f"Private Function {name}(ws As Object) As String", "Dim out As String, v As Variant",
             "On Error Resume Next"]
    for expression in ["ws.UsedRange.Address", *reads]:
        lines += ["Err.Clear", "v = Empty", f"v = {expression}",
                  'If Err.Number <> 0 Then out = out & "E" & Err.Number & ";" Else out = out & Show(v) & ";"']
    return "\n".join([*lines, "On Error GoTo 0", f"{name} = out", "End Function"]) + "\n"


def test_excel_reads_model_authored_row_formats(tmp_path: Path) -> None:
    """Every row, column and sheet format case the model saves reads in Excel as Excel's own file of it does."""
    import json

    harness = pytest.importorskip("pyvbaharness")
    fixture = json.loads((Path(__file__).parent / "fixtures" / "row_formats" / "row_formats.json").read_text())
    cases = [case for case in fixture["cases"] if not case["name"].startswith("unmeasured_")]
    app = ExcelApplication()
    app.add_workbook()
    lines = ["Public Sub Build()", "Dim ws As Object"]
    for index, case in enumerate(cases):
        lines.append("Set ws = ActiveWorkbook.Worksheets(1)" if index == 0 else
                     "Set ws = ActiveWorkbook.Worksheets.Add(After:=ActiveWorkbook.Worksheets"
                     "(ActiveWorkbook.Worksheets.Count))")
        lines += [f'ws.Name = "{case["name"]}"', *case["setup"].splitlines()]
    app.add_module("\n".join([*lines, "ActiveWorkbook.Worksheets(1).Activate", "End Sub"]) + "\n", name="Builder")
    app.run("Build")
    path = tmp_path / "model_row_formats.xlsx"
    app.save(path)
    reader = fixture["helper"] + "".join(_row_format_reads(f"Reads{index}", case["reads"])
                                         for index, case in enumerate(cases))
    reader += ("Public Function Report() As String\n"
               + "".join(f'Report = Report & Reads{index}(ActiveWorkbook.Worksheets({index + 1})) & "|"\n'
                         for index in range(len(cases))) + "End Function\n")
    with harness.ExcelSession(harness.HarnessConfig(lock_wait_s=45.0)) as excel:
        excel.open_document(path)
        result = excel.run_vba(reader, "Report", timeout=240.0)
        assert result.ok, result.message
    for case, answer in zip(cases, str(result.value).split("|")[:-1], strict=True):
        assert answer == case["after"], case["name"]


def test_excel_reads_model_typed_cells(tmp_path: Path) -> None:
    """Cells the model typed read in Excel -- type, number, format, prefix and what shows -- as Excel's own do."""
    import json

    from pyopenvba.apps.excel import _typing

    harness = pytest.importorskip("pyvbaharness")
    folder = Path(__file__).parent / "fixtures" / "typing"
    fixture = json.loads((folder / "typing.json").read_text(encoding="utf-8"))
    lines = ["Public Sub Build()", "Dim c As Object"]
    for row, steps in enumerate(fixture["cases"], start=1):
        lines.append(f"Set c = Cells({row}, 1)")
        lines += [step if step.startswith("c.") else f"c.Value = {step}" for step in steps]
    app = ExcelApplication()
    app.add_workbook()
    app.add_module("\n".join([*lines, "End Sub"]) + "\n", name="Builder")
    with pytest.MonkeyPatch.context() as patch:
        patch.setattr(_typing, "_this_year", lambda: fixture["year"])
        app.run("Build")
    path = tmp_path / "model_typing.xlsx"
    app.save(path)
    count = len(fixture["cases"])
    reader = (
        "Private Function ReadAll(ws As Object) As String\n"
        "Dim out As String, c As Object, i As Long\n"
        f"For i = 1 To {count}\n"
        "Set c = ws.Cells(i, 1)\n"
        'out = out & TypeName(c.Value) & "~" & CStr(c.Value2) & "~" & c.NumberFormat & "~" & '
        'c.PrefixCharacter & "~" & c.Text & "|"\n'
        "Next\nReadAll = out\nEnd Function\n"
        "Public Function Report() As String\n"
        "Dim ours As Object, theirs As Object\n"
        "Application.DisplayAlerts = False\n"
        "Set ours = ActiveWorkbook\n"
        f'Set theirs = Workbooks.Open("{folder / "typing.xlsx"}", ReadOnly:=True)\n'
        'Report = ReadAll(ours.Worksheets(1)) & "<ours|theirs>" & ReadAll(theirs.Worksheets(1))\n'
        "theirs.Close False\nEnd Function\n"
    )
    with harness.ExcelSession(harness.HarnessConfig(lock_wait_s=45.0)) as excel:
        excel.open_document(path)
        result = excel.run_vba(reader, "Report", timeout=240.0)
        assert result.ok, result.message
    ours, theirs = str(result.value).split("<ours|theirs>", 1)
    for steps, mine, excels in zip(fixture["cases"], ours.split("|")[:count], theirs.split("|")[:count], strict=True):
        assert mine == excels, steps


def test_excel_reads_the_filters_this_wrote(tmp_path: Path) -> None:
    """Every kind of filter tests/fixtures/autofilter_file measured, made by the model and opened in Excel."""
    import json

    harness = pytest.importorskip("pyvbaharness")
    record = json.loads((Path(__file__).parent / "fixtures" / "autofilter_file" / "autofilter_file.json")
                        .read_text(encoding="utf-8"))
    build = ["Public Sub Build()", "Dim wb As Object, ws As Object", "Set wb = ActiveWorkbook"]
    bodies: list[str] = []
    for index, case in enumerate(record["cases"]):
        build += ["Set ws = wb.Worksheets(1)" if index == 0 else
                  "Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))",
                  f'ws.Name = "{case["name"]}"', f"Case{index} ws"]
        bodies.append("\n".join([f"Private Sub Case{index}(ws As Object)", "Dim v As Variant",
                                 *(record["table"] + case["setup"]).splitlines(), "End Sub"]) + "\n")
    build += ["End Sub"]
    app = ExcelApplication()
    app.add_workbook()
    app.add_module("\n".join(build) + "\n" + "".join(bodies), name="Builder")
    app.run("Build")
    path = tmp_path / "filters.xlsx"
    app.save(path)
    reader = (record["helper"] + "Public Function Report() As String\nDim ws As Object, out As String\n"
              "For Each ws In ActiveWorkbook.Worksheets\nout = out & ws.Name & \"^\" & Dump(ws) & \"|\"\nNext\n"
              "Report = out\nEnd Function\n")
    with harness.ExcelSession(harness.HarnessConfig(lock_wait_s=45.0)) as excel:
        excel.open_document(path)
        result = excel.run_vba(reader, "Report", timeout=240.0)
        assert result.ok, result.message
    read = dict(one.split("^", 1) for one in str(result.value).split("|") if one)
    for case in record["cases"]:
        assert read[case["name"]] == case["reopened"], case["name"]


def test_excel_reads_the_cell_shifts_this_made(tmp_path: Path) -> None:
    """Cells tests/fixtures/cell_shifts.json shifted, made by the model, saved, and opened in Excel one by one."""
    import json

    harness = pytest.importorskip("pyvbaharness")
    record = json.loads((Path(__file__).parent / "fixtures" / "cell_shifts.json").read_text(encoding="utf-8"))
    chosen = ("del_up_b3b4", "del_up_b2b6", "ins_down_b3c4", "del_left_b3c5", "ins_right_a5", "del_default_b3b4",
              "ins_down_formats", "filter_ins_a3c3", "filter_del_header", "filter_ins_right_a1")
    layouts = [layout for layout in record["layouts"] if layout["name"] in chosen]
    paths: list[Path] = []
    for layout in layouts:
        code = ["Public Sub Build()", "Dim ws As Object, other As Object, r As Long, c As Long",
                "Set ws = ActiveWorkbook.Worksheets(1)", 'ws.Name = "Grid"',
                "Set other = ActiveWorkbook.Worksheets.Add(After:=ws)", 'other.Name = "Other"',
                *record["grid"].splitlines(), *layout["setup"].splitlines(), "End Sub"]
        app = ExcelApplication()
        app.add_workbook()
        app.add_module("\n".join(code) + "\n", name="Builder")
        app.run("Build")
        path = tmp_path / f"{layout['name']}.xlsx"
        app.save(path)
        paths.append(path)
    reader = record["helper"] + 'Public Function Report() As String\nReport = Dump(Worksheets("Grid"))\nEnd Function\n'
    read: dict[str, str] = {}
    with harness.ExcelSession(harness.HarnessConfig(lock_wait_s=45.0)) as excel:
        for layout, path in zip(layouts, paths, strict=True):
            excel.open_document(path)
            result = excel.run_vba(reader, "Report", timeout=120.0)
            assert result.ok, result.message
            read[layout["name"]] = str(result.value)
    for layout in layouts:
        assert read[layout["name"]] == layout["answers"], layout["name"]


def test_excel_reads_the_filtered_edits_this_made(tmp_path: Path) -> None:
    """Every edit tests/fixtures/autofilter_edits.json measured, made by the model, saved and opened in Excel."""
    import json

    from test_excel_autofilter_edits import UNSUPPORTED

    harness = pytest.importorskip("pyvbaharness")
    record = json.loads((Path(__file__).parent / "fixtures" / "autofilter_edits.json").read_text(encoding="utf-8"))
    layouts = [layout for layout in record["layouts"] if layout["name"] not in UNSUPPORTED]
    build = ["Public Sub Build()", "Dim wb As Object, ws As Object", "Set wb = ActiveWorkbook"]
    bodies: list[str] = []
    for index, layout in enumerate(layouts):
        build += ["Set ws = wb.Worksheets(1)" if index == 0 else
                  "Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))",
                  f'ws.Name = "{layout["name"][:28]}"', f"Case{index} ws"]
        bodies.append("\n".join([f"Private Sub Case{index}(ws As Object)", "Dim v As Variant, dest As Object",
                                 "On Error Resume Next", *(record["table"] + layout["setup"]).splitlines(),
                                 "End Sub"]) + "\n")
    build += ["End Sub"]
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(record["helper"] + "\n".join(build) + "\n" + "".join(bodies), name="Builder")
    app.run("Build")
    path = tmp_path / "filtered_edits.xlsx"
    app.save(path)
    reader = (record["helper"] + "Public Function Report() As String\nDim ws As Object, out As String\n"
              "For Each ws In ActiveWorkbook.Worksheets\nout = out & ws.Name & \"^\" & Dump(ws) & \"|\"\nNext\n"
              "Report = out\nEnd Function\n")
    with harness.ExcelSession(harness.HarnessConfig(lock_wait_s=45.0)) as excel:
        excel.open_document(path)
        result = excel.run_vba(reader, "Report", timeout=240.0)
        assert result.ok, result.message
    read = dict(one.split("^", 1) for one in str(result.value).split("|") if one)
    for layout in layouts:
        # What Excel read after the edit, past the error and the value the setup left.
        assert read[layout["name"][:28]] == "hidden=" + layout["answers"].split(";hidden=", 1)[1], layout["name"]
