"""What a macro reads of a workbook another tool wrote, where nearly everything else was measured on files Excel or the model
wrote. One book, made the same way by openpyxl and by XlsxWriter, holds cells, styles, names, a table, conditional formats,
validation, a note, a hyperlink, a chart and a picture. Excel and the model each read every member, one by one, typed and late
bound, recording its type and value or its error. Then each makes an edit and saves, and Excel reads both results.

Neither tool is a dependency of the model: XlsxWriter is only needed to remake the fixture.

    PYTHONPATH=<folder holding xlsxwriter> python scripts/measure_tool_books.py

writes tests/fixtures/tool_books/, which tests/test_excel_tool_books.py replays.
"""
from __future__ import annotations

import datetime as dt
import json
import shutil
import struct
import sys
import tempfile
import zlib
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))
OUT = ROOT / "tests/fixtures/tool_books"

#: What is read, by its text, of ws (the sheet Data), other (the sheet Other) and wb.
READS: list[str] = [
    # Cells
    'ws.Range("A1").Value', 'ws.Range("A2").Value', 'ws.Range("B2").Value', 'ws.Range("B2").Value2', 'ws.Range("C2").Value', 'ws.Range("C2").Value2',
    'ws.Range("C2").Text', 'ws.Range("D2").Value', 'ws.Range("E2").Value', 'ws.Range("E2").Formula', 'ws.Range("E2").FormulaR1C1', 'ws.Range("E2").Text',
    'ws.Range("E2").HasFormula', 'ws.Range("E7").Value', 'ws.Range("E7").Formula', 'ws.Range("E8").Value', 'ws.Range("E8").Text', 'ws.Range("E9").Value',
    'ws.Range("E10").Value', 'ws.Range("Z99").Value', 'ws.UsedRange.Address', 'ws.Cells(ws.Rows.Count, 1).End(-4162).Row', 'wb.Worksheets.Count', 'wb.Sheets(2).Name',
    'ws.Range("F2").Value', 'ws.Range("F2").Text', 'ws.Range("F3").Text', 'ws.Range("F4").Value', 'ws.Range("F4").Text',
    # Styles
    'ws.Range("A1").Font.Bold', 'ws.Range("A1").Font.Name', 'ws.Range("A1").Font.Size', 'ws.Range("A2").Font.Italic', 'ws.Range("A2").Font.Color', 'ws.Range("A2").Font.ColorIndex',
    'ws.Range("A3").Font.Name', 'ws.Range("A3").Font.Size', 'ws.Range("A3").Font.Bold', 'ws.Range("A3").Font.Color', 'ws.Range("A3").Interior.Pattern', 'ws.Range("A3").Interior.ColorIndex',
    'ws.Range("B2").Interior.Color', 'ws.Range("B2").Interior.Pattern', 'ws.Range("B3").Borders(9).LineStyle', 'ws.Range("B3").Borders(9).Weight', 'ws.Range("B3").Borders(7).LineStyle',
    'ws.Range("C3").HorizontalAlignment', 'ws.Range("C3").VerticalAlignment', 'ws.Range("C4").WrapText', 'ws.Range("B2").NumberFormat', 'ws.Range("C2").NumberFormat', 'ws.Range("F2").NumberFormat',
    'ws.Range("F3").NumberFormat', 'ws.Range("A1").Style.Name', 'ws.Range("B3").Locked', 'wb.Styles.Count', 'wb.Styles("Normal").Font.Name', 'wb.Styles("Normal").Font.Size',
    # Layout
    'ws.Range("H1").MergeCells', 'ws.Range("H1").MergeArea.Address', 'ws.Rows(3).Hidden', 'ws.Rows(7).Hidden', 'ws.Columns(11).Hidden', 'ws.Rows(3).RowHeight', 'ws.Rows(4).RowHeight',
    'ws.Columns(1).ColumnWidth', 'ws.Columns(2).ColumnWidth', 'ws.StandardHeight', 'ws.StandardWidth', 'other.Tab.Color', 'ws.Visible', 'wb.Worksheets("Spare").Visible', 'ws.ProtectContents',
    # Names
    'wb.Names.Count', 'wb.Names("Total").RefersTo', 'wb.Names("Total").Visible', 'other.Names.Count', 'other.Names("Own").RefersTo', 'ws.Range("Total").Address', 'ws.PageSetup.PrintArea', 'ws.PageSetup.Orientation',
    # Table
    'other.ListObjects.Count', 'other.ListObjects(1).Name', 'other.ListObjects(1).Range.Address', 'other.ListObjects(1).ListColumns.Count', 'other.ListObjects(1).ListColumns(2).Name',
    'other.ListObjects(1).ShowTotals', 'other.ListObjects(1).TableStyle', 'other.ListObjects(1).ShowAutoFilter', 'other.ListObjects(1).ListRows.Count', 'other.ListObjects(1).DataBodyRange.Address',
    # Conditional formats
    'ws.Cells.FormatConditions.Count', 'ws.Range("B2").FormatConditions.Count', 'ws.Range("B2").FormatConditions(1).Type', 'ws.Range("B2").FormatConditions(1).Operator', 'ws.Range("B2").FormatConditions(1).Formula1',
    'ws.Range("B2").FormatConditions(1).Interior.Color', 'ws.Range("B2").FormatConditions(1).AppliesTo.Address', 'ws.Range("L2").FormatConditions.Count', 'ws.Range("L2").FormatConditions(1).Type', 'ws.Range("L2").FormatConditions(1).ColorScaleCriteria.Count',
    'ws.Range("M2").FormatConditions(1).Type', 'ws.Range("N2").FormatConditions(1).Type', 'ws.Range("N2").FormatConditions(1).IconSet.ID',
    # Validation
    'ws.Range("G2").Validation.Type', 'ws.Range("G2").Validation.Formula1', 'ws.Range("G2").Validation.InCellDropdown', 'ws.Range("G2").Validation.ErrorMessage', 'ws.Range("G2").Validation.ShowError', 'ws.Range("G3").Validation.Type',
    'ws.Range("G3").Validation.Operator', 'ws.Range("G3").Validation.Formula1', 'ws.Range("G3").Validation.Formula2',
    # Notes and links
    'ws.Comments.Count', 'ws.Range("A4").Comment.Text', 'ws.Range("A4").Comment.Author', 'ws.Range("A4").Comment.Visible', 'ws.Range("A5").Hyperlinks.Count', 'ws.Range("A5").Hyperlinks(1).Address', 'ws.Range("A5").Hyperlinks(1).TextToDisplay', 'ws.Hyperlinks.Count',
    # Drawings
    'ws.Shapes.Count', 'ws.ChartObjects.Count', 'ws.ChartObjects(1).Chart.ChartType', 'ws.ChartObjects(1).Chart.HasTitle', 'ws.ChartObjects(1).Chart.ChartTitle.Text', 'ws.ChartObjects(1).Chart.HasLegend',
    'ws.ChartObjects(1).Chart.SeriesCollection.Count', 'ws.ChartObjects(1).Chart.SeriesCollection(1).Formula', 'ws.ChartObjects(1).Chart.SeriesCollection(1).Name', 'ws.ChartObjects(1).TopLeftCell.Address',
    'ws.Shapes(2).Type', 'ws.Pictures.Count', 'ws.Shapes(2).TopLeftCell.Address',
    # Window
    'wb.ActiveSheet.Name', 'wb.Windows(1).FreezePanes', 'wb.Windows(1).SplitRow', 'wb.Windows(1).Zoom', 'wb.Windows(1).DisplayGridlines', 'wb.Date1904', 'Application.Calculation',
    # The file
    'wb.BuiltinDocumentProperties("Title").Value', 'wb.BuiltinDocumentProperties("Author").Value', 'wb.FileFormat', 'wb.HasVBProject', 'wb.Saved',
]
#: What is done before a save, both ways.
EDITS: dict[str, str] = {"far_cell": 'ws.Range("Z99").Value = 7', "changes_a_read_cell": 'ws.Range("B2").Value = 100', "row_inserted": "ws.Rows(1).Insert"}
SETUP = '''    Set wb = Workbooks.Open(file, UpdateLinks:=0)
    Set ws = wb.Worksheets("Data")
    Set other = wb.Worksheets("Other")
'''
ONE = '''Private Function Shown(ByRef v As Variant) As String
    If IsObject(v) Then
        Shown = "Object"
    ElseIf IsError(v) Then
        Shown = "Error:" & CStr(v)
    ElseIf IsNull(v) Then
        Shown = "Null"
    ElseIf IsEmpty(v) Then
        Shown = "Empty"
    Else
        Shown = TypeName(v) & ":" & CStr(v)
    End If
End Function
Function Go(ByVal file As String) As String
    Dim wb As {book}, ws As {sheet}, other As {sheet}, v As Variant, s As String
    Application.DisplayAlerts = False
{setup}{reads}    wb.Close False
    Go = Mid$(s, 2)
End Function
'''
READ_ONE = '''    On Error Resume Next
    v = Empty
    v = {read}
    If Err.Number <> 0 Then
        s = s & vbTab & "E" & Err.Number
        Err.Clear
    Else
        On Error GoTo 0
        s = s & vbTab & Shown(v)
    End If
    On Error GoTo 0
'''
EDIT = '''Sub Act(ByVal file As String)
    Dim wb As Workbook, ws As Worksheet, other As Worksheet
    Application.DisplayAlerts = False
{setup}    {action}
    wb.Save
    wb.Close False
End Sub
'''
BINDINGS = {"typed": ("Workbook", "Worksheet"), "late": ("Object", "Object")}


def source(read_list: list[str], binding: str) -> str:
    book, sheet = BINDINGS[binding]
    return ONE.format(book=book, sheet=sheet, setup=SETUP, reads="".join(READ_ONE.format(read=read) for read in read_list))


def png(file: Path) -> None:
    def chunk(kind: bytes, data: bytes) -> bytes:
        return struct.pack(">I", len(data)) + kind + data + struct.pack(">I", zlib.crc32(kind + data))
    rows = (b"\x00" + b"\x10\x80\xf0" * 16) * 12
    file.write_bytes(b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", 16, 12, 8, 2, 0, 0, 0)) + chunk(b"IDAT", zlib.compress(rows)) + chunk(b"IEND", b""))


DAYS = [dt.datetime(2024, 1, 15), dt.datetime(2024, 2, 20), dt.datetime(2024, 3, 25), dt.datetime(2024, 4, 30), dt.datetime(2024, 5, 5)]


def by_openpyxl(file: Path, picture: Path) -> None:
    from openpyxl import Workbook
    from openpyxl.chart import BarChart, Reference
    from openpyxl.comments import Comment
    from openpyxl.drawing.image import Image
    from openpyxl.formatting.rule import CellIsRule, ColorScaleRule, DataBarRule, IconSetRule
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.workbook.defined_name import DefinedName
    from openpyxl.worksheet.datavalidation import DataValidation
    from openpyxl.worksheet.table import Table, TableStyleInfo

    wb = Workbook()
    ws = wb.active
    ws.title = "Data"
    other = wb.create_sheet("Other")
    spare = wb.create_sheet("Spare")
    spare.sheet_state = "hidden"
    for column, text in enumerate(("Name", "Value", "When", "Flag", "Calc"), 1):
        ws.cell(1, column, text)
    for row in range(2, 7):
        ws.cell(row, 1, "abcde"[row - 2])
        ws.cell(row, 2, row * 10)
        when = ws.cell(row, 3, DAYS[row - 2])
        when.number_format = "yyyy-mm-dd"
        ws.cell(row, 4, row % 2 == 0)
        ws.cell(row, 5, f"=B{row}*2")
    ws["E7"] = "=SUM(B2:B6)"
    ws["E8"] = "=1/0"
    ws["E9"] = "=Other!A1+1"
    ws["E10"] = '="x"&"y"'
    ws["F2"] = 1.5
    ws["F2"].number_format = "0.00"
    ws["F3"] = 0.25
    ws["F3"].number_format = "0.0%"
    ws["F4"] = dt.datetime(2024, 6, 1, 13, 30)
    ws["F4"].number_format = "m/d/yy h:mm"
    ws["A1"].font = Font(name="Calibri", size=11, bold=True)
    ws["A2"].font = Font(name="Calibri", size=11, italic=True, color="FFFF0000")
    ws["B2"].fill = PatternFill("solid", fgColor="FFFFFF00")
    ws["B3"].border = Border(bottom=Side(style="thin"))
    ws["C3"].alignment = Alignment(horizontal="center", vertical="top")
    ws["C4"].alignment = Alignment(wrap_text=True)
    ws.merge_cells("H1:I1")
    ws["H1"] = "Merged"
    ws.row_dimensions[3].height = 30
    ws.row_dimensions[7].hidden = True
    ws.column_dimensions["A"].width = 20
    ws.column_dimensions["K"].hidden = True
    ws.freeze_panes = "A2"
    other.sheet_properties.tabColor = "FF00FF00"
    other["A1"] = 5
    wb.defined_names["Total"] = DefinedName("Total", attr_text="Data!$B$2:$B$6")
    other.defined_names["Own"] = DefinedName("Own", attr_text="Other!$A$1")
    other["D1"], other["E1"] = "Key", "Val"
    for row in range(2, 5):
        other.cell(row, 4, "pqr"[row - 2])
        other.cell(row, 5, row)
    table = Table(displayName="Tbl", ref="D1:E4")
    table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    other.add_table(table)
    ws.conditional_formatting.add("B2:B6", CellIsRule(operator="greaterThan", formula=["30"], fill=PatternFill(bgColor="FFFF0000", fill_type="solid")))
    ws.conditional_formatting.add("L2:L6", ColorScaleRule(start_type="min", start_color="FFFF0000", end_type="max", end_color="FF00FF00"))
    ws.conditional_formatting.add("M2:M6", DataBarRule(start_type="min", end_type="max", color="FF0000FF"))
    ws.conditional_formatting.add("N2:N6", IconSetRule("3TrafficLights1", "percent", [0, 33, 67]))
    listed = DataValidation(type="list", formula1='"x,y,z"', allow_blank=True, showErrorMessage=True, error="Pick one")
    listed.add("G2")
    ws.add_data_validation(listed)
    whole = DataValidation(type="whole", operator="between", formula1="1", formula2="9", showErrorMessage=True)
    whole.add("G3")
    ws.add_data_validation(whole)
    ws["A4"].comment = Comment("A note", "Author")
    ws["A5"].hyperlink = "http://example.com/"
    chart = BarChart()
    chart.title = "Sales"
    chart.add_data(Reference(ws, min_col=2, min_row=1, max_row=6), titles_from_data=True)
    chart.set_categories(Reference(ws, min_col=1, min_row=2, max_row=6))
    ws.add_chart(chart, "P2")
    ws.add_image(Image(str(picture)), "P20")
    ws.page_setup.orientation = "landscape"
    ws.print_area = "A1:E7"
    wb.properties.title = "Made"
    wb.properties.creator = "Maker"
    wb.save(file)


def by_xlsxwriter(file: Path, picture: Path) -> None:
    import xlsxwriter

    wb = xlsxwriter.Workbook(str(file))
    ws = wb.add_worksheet("Data")
    other = wb.add_worksheet("Other")
    spare = wb.add_worksheet("Spare")
    spare.hide()
    ws.write_row("A1", ("Name", "Value", "When", "Flag", "Calc"))
    date = wb.add_format({"num_format": "yyyy-mm-dd"})
    for row in range(2, 7):
        ws.write_string(row - 1, 0, "abcde"[row - 2])
        ws.write_number(row - 1, 1, row * 10)
        ws.write_datetime(row - 1, 2, DAYS[row - 2], date)
        ws.write_boolean(row - 1, 3, row % 2 == 0)
        ws.write_formula(row - 1, 4, f"=B{row}*2")
    ws.write_formula("E7", "=SUM(B2:B6)")
    ws.write_formula("E8", "=1/0")
    ws.write_formula("E9", "=Other!A1+1")
    ws.write_formula("E10", '="x"&"y"')
    ws.write_number("F2", 1.5, wb.add_format({"num_format": "0.00"}))
    ws.write_number("F3", 0.25, wb.add_format({"num_format": "0.0%"}))
    ws.write_datetime("F4", dt.datetime(2024, 6, 1, 13, 30), wb.add_format({"num_format": "m/d/yy h:mm"}))
    ws.write_string("A1", "Name", wb.add_format({"bold": True}))
    ws.write_string("A2", "a", wb.add_format({"italic": True, "font_color": "#FF0000"}))
    ws.write_number("B2", 20, wb.add_format({"bg_color": "#FFFF00", "pattern": 1}))
    ws.write_number("B3", 30, wb.add_format({"bottom": 1}))
    ws.write_datetime("C3", DAYS[1], wb.add_format({"align": "center", "valign": "top", "num_format": "yyyy-mm-dd"}))
    ws.write_datetime("C4", DAYS[2], wb.add_format({"text_wrap": True, "num_format": "yyyy-mm-dd"}))
    ws.merge_range("H1:I1", "Merged")
    ws.set_row(2, 30)
    ws.set_row(6, None, None, {"hidden": True})
    ws.set_column("A:A", 20)
    ws.set_column("K:K", None, None, {"hidden": True})
    ws.freeze_panes(1, 0)
    other.set_tab_color("#00FF00")
    other.write_number("A1", 5)
    wb.define_name("Total", "=Data!$B$2:$B$6")
    wb.define_name("Other!Own", "=Other!$A$1")
    other.add_table("D1:E4", {"name": "Tbl", "style": "Table Style Medium 2", "columns": [{"header": "Key"}, {"header": "Val"}],
                              "data": [[letter, value] for letter, value in zip("pqr", (2, 3, 4))]})
    ws.conditional_format("B2:B6", {"type": "cell", "criteria": ">", "value": 30, "format": wb.add_format({"bg_color": "#FF0000"})})
    ws.conditional_format("L2:L6", {"type": "2_color_scale", "min_color": "#FF0000", "max_color": "#00FF00"})
    ws.conditional_format("M2:M6", {"type": "data_bar", "bar_color": "#0000FF"})
    ws.conditional_format("N2:N6", {"type": "icon_set", "icon_style": "3_traffic_lights"})
    ws.data_validation("G2", {"validate": "list", "source": ["x", "y", "z"], "error_message": "Pick one"})
    ws.data_validation("G3", {"validate": "integer", "criteria": "between", "minimum": 1, "maximum": 9})
    ws.write_comment("A4", "A note", {"author": "Author"})
    ws.write_url("A5", "http://example.com/", string="e")
    chart = wb.add_chart({"type": "column"})
    chart.add_series({"name": "=Data!$B$1", "categories": "=Data!$A$2:$A$6", "values": "=Data!$B$2:$B$6"})
    chart.set_title({"name": "Sales"})
    ws.insert_chart("P2", chart)
    ws.insert_image("P20", str(picture))
    ws.set_landscape()
    ws.print_area("A1:E7")
    wb.set_properties({"title": "Made", "author": "Maker"})
    wb.close()


def model_reads(file: Path, read_list: list[str], binding: str) -> list[str]:
    """Each read alone, in a new model, so a report stops that read and no other."""
    from pyopenvba.apps.excel import ExcelApplication
    from pyopenvba.exceptions import VBAUnsupportedError

    answers: list[str] = []
    for read in read_list:
        app = ExcelApplication(sandbox=False)
        app.add_workbook()
        app.add_module(source([read], binding), name="M")
        try:
            answers.append(str(app.run("Go", file.as_posix())))
        except VBAUnsupportedError as failure:
            answers.append(f"reports: {str(failure)[:90]}")
        except Exception as failure:  # noqa: BLE001 - what the model does with such a file is what is measured
            answers.append(f"crash: {type(failure).__name__}: {str(failure)[:90]}")
    return answers


def tally(excel: list[str], model: list[str]) -> dict[str, int]:
    found = {"same": 0, "differ": 0, "reports": 0, "crash": 0}
    for one, other in zip(excel, model, strict=True):
        found["reports" if other.startswith("reports:") else "crash" if other.startswith("crash:") else "same" if one == other else "differ"] += 1
    return found


def model_edit(file: Path, action: str) -> str:
    from pyopenvba.apps.excel import ExcelApplication
    from pyopenvba.exceptions import VBAUnsupportedError

    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(EDIT.format(setup=SETUP, action=action), name="M")
    try:
        app.run("Act", file.as_posix())
    except VBAUnsupportedError as failure:
        return f"reports: {str(failure)[:100]}"
    except Exception as failure:  # noqa: BLE001 - as above
        return f"crash: {type(failure).__name__}: {str(failure)[:100]}"
    return ""


def main() -> None:
    from pyvbaharness import ExcelSession, HarnessConfig

    OUT.mkdir(parents=True, exist_ok=True)
    folder = Path(tempfile.mkdtemp(prefix="pyopenvba_tools_"))
    picture = folder / "p.png"
    png(picture)
    record: dict[str, object] = {"reads": READS, "edits": EDITS, "setup": SETUP}
    tools: dict[str, object] = {}
    try:
        for tool, make in (("openpyxl", by_openpyxl), ("xlsxwriter", by_xlsxwriter)):
            file = OUT / f"{tool}.xlsx"
            make(file, picture)
            row: dict[str, object] = {}
            for binding in BINDINGS:
                copy = folder / f"{tool}_{binding}.xlsx"
                shutil.copy(file, copy)
                with ExcelSession(HarnessConfig(lock_wait_s=900.0)) as excel:
                    excel.new_document()
                    read = excel.run_vba(source(READS, binding), "Go", (str(copy),), timeout=120.0, line_numbers=False)
                assert read.ok, (tool, binding, read.outcome, read.message)
                answers = str(read.value).split("\t")
                assert len(answers) == len(READS), (len(answers), len(READS))
                mine = model_reads(copy, READS, binding)
                row[binding] = {"excel": answers, "model": mine}
                print(tool, binding, tally(answers, mine), flush=True)
            edits: dict[str, object] = {}
            for name, action in EDITS.items():
                theirs, ours = folder / f"{tool}_{name}_excel.xlsx", folder / f"{tool}_{name}_model.xlsx"
                shutil.copy(file, theirs)
                shutil.copy(file, ours)
                reported = model_edit(ours, action)
                with ExcelSession(HarnessConfig(lock_wait_s=900.0)) as excel:
                    excel.new_document()
                    acted = excel.run_vba(EDIT.format(setup=SETUP, action=action), "Act", (str(theirs),), timeout=120.0)
                    assert acted.ok, (tool, name, acted.message)
                    after = excel.run_vba(source(READS, "typed"), "Go", (str(theirs),), timeout=120.0, line_numbers=False)
                    mine_after = None
                    if not reported:
                        mine_after = excel.run_vba(source(READS, "typed"), "Go", (str(ours),), timeout=120.0, line_numbers=False)
                entry: dict[str, object] = {"excel": str(after.value).split("\t"), "reported": reported}
                if mine_after is not None:
                    entry["model"] = str(mine_after.value).split("\t") if mine_after.ok else [f"failed {mine_after.outcome} {mine_after.message[:80]}"]
                edits[name] = entry
                differ = sum(a != b for a, b in zip(entry["excel"], entry.get("model", []))) if len(entry.get("model", [])) == len(READS) else "?"
                print(tool, name, "reported" if reported else "saved", reported[:90], "reads differing", differ, flush=True)
            row["edits"] = edits
            tools[tool] = row
    finally:
        shutil.rmtree(folder, ignore_errors=True)
    record["tools"] = tools
    record["measured_at"] = datetime.now(timezone.utc).isoformat()
    (OUT / "tool_books.json").write_text(json.dumps(record, indent=1) + "\n", encoding="utf-8")


if __name__ == "__main__":
    main()
