"""How Excel fits a pivot table's columns each time it lays the table out, measured in live Excel.

Five parts, each in a session of its own:

- styles: every built-in table and pivot table style, and for each of its
  elements whether it draws its font bold, as TableStyleElements(...).Font
  reads it.
- compact: the width Excel fits a compact table's label column to, for an
  item of each of ten texts at every depth of one to four row fields, with
  and without the +/- buttons and at compact indents 0 to 3, every other
  cell made too small to matter.
- subtotals: the same for a subtotal's label below its items.
- layouts: tables in outline and tabular form with column fields, with and
  without grand totals, and the widths of their columns.
- steps: tables edited a step at a time, laying them out, renaming,
  restyling, refreshing and the rest, and their columns' widths after each.

    python scripts/measure_pivot_column_fit.py [part ...]

writes tests/fixtures/pivot_column_fit/pivot_column_fit.json, which
tests/test_excel_pivot_column_fit.py replays. Widths are pixels on a 96-DPI
display, each column's Width times 4/3.
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "tests" / "fixtures" / "pivot_column_fit"

TEXTS = ["Mx", "Kiwi fruit", "Wm Wm", "Apples and pears", "Mountains", "Wwwwwwwwwwwwww", "Quick brown fox jumps",
         "Lorem ipsum dolor sit amet", "Mmmmmmmmmmmmmmmmmmmm", "Extraordinarily long product name"]

STYLES = '''
Public Function Styles() As String
    Dim wb As Object, ts As Object, out As String, e As Long, el As Object, v As Variant
    Set wb = Workbooks.Add(xlWBATWorksheet)
    For Each ts In wb.TableStyles
        If ts.BuiltIn Then
            out = out & ts.Name & "~" & IIf(ts.ShowAsAvailablePivotTableStyle, "p", "") & _
                IIf(ts.ShowAsAvailableTableStyle, "t", "") & ":"
            On Error Resume Next
            For e = 0 To 27
                Err.Clear
                Set el = Nothing
                Set el = ts.TableStyleElements(e)
                If Err.Number = 0 And Not el Is Nothing Then
                    If el.HasFormat Then
                        v = el.Font.Bold
                        If Not IsNull(v) Then If v Then out = out & e & ","
                    End If
                End If
            Next
            On Error GoTo 0
            out = out & "|"
        End If
    Next
    wb.Close False
    Styles = out
End Function
'''

#: A table of one to four row fields L0 to L3 over one record, its row header R, no grand totals and a data field S,
#: with the text measured at one depth and single letters at the others.
GRID = '''
Private Function Made(wb As Object, depth As Long, fields As Long, drill As Boolean, indent As Long, text As String, _
        totals As Boolean) As Object
    Dim data As Object, ws As Object, pt As Object, f As Long
    Set data = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    data.Range("A1:E1").Value = Array("L0", "L1", "L2", "L3", "Amount")
    For f = 0 To 3
        data.Cells(2, f + 1).Value = IIf(f = depth, text, Chr(97 + f))
    Next
    data.Cells(2, 5).Value = 1
    Set ws = wb.Worksheets.Add(After:=data)
    Set pt = wb.PivotCaches.Create(xlDatabase, data.Name & "!R1C1:R2C5").CreatePivotTable(ws.Range("A3"))
    pt.CompactLayoutRowHeader = "R"
    pt.ColumnGrand = False
    pt.RowGrand = False
    pt.ShowDrillIndicators = drill
    pt.CompactRowIndent = indent
    For f = 0 To fields - 1
        pt.PivotFields("L" & f).Orientation = xlRowField
        If totals And f = depth Then
            pt.PivotFields("L" & f).Subtotals(1) = True
            pt.PivotFields("L" & f).LayoutSubtotalLocation = xlAtBottom
        Else
            pt.PivotFields("L" & f).Subtotals(1) = False
        End If
    Next
    pt.AddDataField pt.PivotFields("Amount"), "S", xlSum
    Set Made = ws
End Function

Private Function Fit(wb As Object, depth As Long, fields As Long, drill As Boolean, indent As Long, text As String, _
        totals As Boolean) As String
    Dim ws As Object
    Set ws = Made(wb, depth, fields, drill, indent, text, totals)
    Fit = Round(ws.Columns(1).Width * 4 / 3)
    Application.DisplayAlerts = False
    wb.Worksheets(ws.Index - 1).Delete
    ws.Delete
End Function

Public Function Grid(t As Long, text As String, totals As Boolean) As String
    Dim wb As Object, out As String, fields As Long, depth As Long, d As Long, indent As Long, last As Long
    Application.DisplayAlerts = False
    Set wb = Workbooks.Add(xlWBATWorksheet)
    For fields = 1 To 4
        last = IIf(totals, fields - 2, fields - 1)
        For depth = 0 To last
            For d = 0 To 1
                For indent = 0 To 3
                    out = out & t & ":" & fields & ":" & depth & ":" & d & ":" & indent & "=" & _
                        Fit(wb, depth, fields, d = 1, indent, text, totals) & ";"
                Next
            Next
        Next
    Next
    wb.Close False
    Grid = out
End Function
'''

#: The source the layouts and steps tables draw on: regions, products with one long name, channels and amounts.
SOURCE = '''
Private Function Book() As Object
    Dim wb As Object, data As Object, r As Long
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set data = wb.Worksheets(1)
    data.Name = "data"
    data.Range("A1:E1").Value = Array("Region", "Product", "Channel", "Amount", "Units")
    For r = 0 To 11
        data.Cells(r + 2, 1).Value = Array("East", "North", "South", "West")(r Mod 4)
        data.Cells(r + 2, 2).Value = Array("Apple", "Banana", "Extraordinarily long product name")(r Mod 3)
        data.Cells(r + 2, 3).Value = Array("Online", "Store")(r Mod 2)
        data.Cells(r + 2, 4).Value = (r + 1) * 10
        data.Cells(r + 2, 5).Value = r + 1
    Next
    Set Book = wb
End Function

Private Function Widths(ws As Object) As String
    Dim c As Long, out As String
    For c = 1 To 14
        out = out & Round(ws.Columns(c).Width * 4 / 3) & ","
    Next
    Widths = out
End Function

Private Function Run(name As String, rows As String, columns As String, step As Long) As String
    Dim wb As Object, ws As Object, pt As Object, f As Variant, out As String
    Set wb = Book()
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(1))
    ws.Name = "report"
    Set pt = wb.PivotCaches.Create(xlDatabase, "data!R1C1:R13C5").CreatePivotTable(ws.Range("A3"), "Report")
    If Len(rows) > 0 Then
        For Each f In Split(rows, ",")
            pt.PivotFields(f).Orientation = xlRowField
        Next
    End If
    If Len(columns) > 0 Then
        For Each f In Split(columns, ",")
            pt.PivotFields(f).Orientation = xlColumnField
        Next
    End If
    pt.AddDataField pt.PivotFields("Amount"), "Sum of Amount", xlSum
    out = "w=" & Widths(ws) & " "
    On Error Resume Next
    Err.Clear
    Select Case step
{steps}
    End Select
    If Err.Number <> 0 Then out = out & "E" & Err.Number & " "
    On Error GoTo 0
    out = out & "w=" & Widths(ws)
    wb.Close False
    Run = name & "^" & out & "|"
End Function

Public Function Scenarios() As String
    Dim out As String
    Application.DisplayAlerts = False
{calls}
    Scenarios = out
End Function
'''

#: name -> (row fields, column fields, what follows, W marking where the widths are read again).
LAYOUTS: dict[str, tuple[str, str, str]] = {
    "outline_one_column": ("Region", "Product", "pt.RowAxisLayout xlOutlineRow"),
    "tabular_one_column": ("Region", "Product", "pt.RowAxisLayout xlTabularRow"),
    "tabular_one_column_no_grand": ("Region", "Product", "pt.RowAxisLayout xlTabularRow\npt.RowGrand = False"),
    "tabular_one_column_no_column_grand": ("Region", "Product", "pt.RowAxisLayout xlTabularRow\npt.ColumnGrand = False"),
    "tabular_one_column_long_caption": ("Region", "Channel", 'pt.RowAxisLayout xlTabularRow\n'
                                        'pt.PivotFields("Channel").Caption = "A channel caption longer than its items"'),
    "tabular_two_columns": ("Region", "Product,Channel", "pt.RowAxisLayout xlTabularRow"),
    "tabular_two_columns_no_grand": ("Region", "Product,Channel", "pt.RowAxisLayout xlTabularRow\npt.RowGrand = False"),
    "outline_two_columns": ("Region", "Product,Channel", "pt.RowAxisLayout xlOutlineRow"),
    "tabular_two_columns_no_subtotals": ("Region", "Product,Channel", 'pt.RowAxisLayout xlTabularRow\n'
                                         'pt.PivotFields("Product").Subtotals(1) = False'),
    "tabular_inner_long": ("Region", "Channel,Product", "pt.RowAxisLayout xlTabularRow"),
    "tabular_column_only": ("", "Product", "pt.RowAxisLayout xlTabularRow"),
    "tabular_column_only_no_grand": ("", "Product", "pt.RowAxisLayout xlTabularRow\npt.RowGrand = False"),
    "tabular_two_rows": ("Region,Channel", "Product", "pt.RowAxisLayout xlTabularRow"),
    "compact_two_columns": ("Region", "Product,Channel", ""),
    "compact_two_columns_no_drill": ("Region", "Product,Channel", "pt.ShowDrillIndicators = False"),
    "column_form_outline": ("Region", "Product", 'pt.PivotFields("Product").LayoutForm = xlOutline'),
    "column_compact_off": ("Region", "Product", 'pt.PivotFields("Product").LayoutCompactRow = False'),
}

STEPS: dict[str, tuple[str, str, str]] = {
    "subtotals_bottom": ("Region,Product", "", 'pt.PivotFields("Region").LayoutSubtotalLocation = xlAtBottom'),
    "inner_subtotals": ("Region,Product", "", 'pt.PivotFields("Product").Subtotals = Array(False, True, True, False, '
                        'False, False, False, False, False, False, False, False)'),
    "grand_caption": ("Region", "", 'pt.GrandTotalName = "A grand total name that is longer"'),
    "style_medium4": ("Region", "Product", 'pt.TableStyle2 = "PivotStyleMedium4"'),
    "style_dark1": ("Region,Product", "", 'pt.TableStyle2 = "PivotStyleDark1"'),
    "style_table": ("Region", "Product", 'pt.TableStyle2 = "TableStyleMedium2"'),
    "flags_off": ("Region", "Product", "pt.ShowTableStyleColumnHeaders = False\nW\npt.ShowTableStyleRowHeaders = False"),
    "flag_last_column": ("Region", "Product", 'pt.TableStyle2 = "PivotStyleMedium22"\nW\n'
                         "pt.ShowTableStyleLastColumn = False"),
    "autoformat_off": ("Region", "", 'pt.HasAutoFormat = False\nW\npt.PivotFields("Product").Orientation = xlRowField'),
    "manual_width": ("Region", "", "ws.Columns(1).ColumnWidth = 40\nW\npt.PivotFields(\"Product\").Orientation = "
                     "xlColumnField"),
    "outside": ("Region", "", 'ws.Range("A40").Value = "Text outside the table that is very long indeed"\nW\n'
                'pt.PivotFields("Product").Orientation = xlColumnField'),
    "page_field": ("Region", "", 'pt.PivotFields("Product").Orientation = xlPageField\nW\n'
                   'pt.PivotFields("Product").CurrentPage = "Extraordinarily long product name"'),
    "page_multiple": ("Region", "", 'pt.PivotFields("Product").Orientation = xlPageField\n'
                      'pt.PivotFields("Product").EnableMultiplePageItems = True\n'
                      'pt.PivotFields("Product").PivotItems("Apple").Visible = False'),
    "values_rows": ("Region", "", 'pt.AddDataField pt.PivotFields("Units"), "Sum of Units", xlSum\nW\n'
                    "pt.DataPivotField.Orientation = xlRowField"),
    "values_columns": ("Region", "Product", 'pt.AddDataField pt.PivotFields("Units"), "Sum of Units", xlSum'),
    "values_row_shown": ("Region", "", 'pt.AddDataField pt.PivotFields("Units"), "Sum of Units", xlSum\n'
                         "pt.ShowValuesRow = True"),
    "number_format": ("Region", "", 'pt.DataFields(1).NumberFormat = "#,##0.000000000000"'),
    "item_caption": ("Region", "", 'pt.PivotFields("Region").PivotItems("East").Caption = "A caption much longer than '
                     'the rest of them"'),
    "refresh": ("Region", "", 'wb.Worksheets("data").Range("A2").Value = "A region name much longer than any other"\n'
                "pt.PivotCache.Refresh"),
    "collapse": ("Region,Product", "", 'pt.PivotFields("Region").ShowDetail = False'),
    "tabular_outer_kept": ("Region,Product", "", "pt.RowAxisLayout xlTabularRow\nW\n"
                           'pt.PivotFields("Region").PivotItems("East").Caption = "A caption much longer than the rest '
                           'of them"'),
    "compact_mixed": ("Region,Product", "", 'pt.PivotFields("Region").LayoutForm = xlTabular'),
    "drill_off": ("Region,Product", "", "pt.ShowDrillIndicators = False"),
    "indent_two": ("Region,Product", "", "pt.CompactRowIndent = 2"),
    "headers_off": ("Region", "Product", "pt.DisplayFieldCaptions = False"),
    "no_rows": ("", "Product", ""),
    "no_columns_no_rows": ("", "", 'pt.AddDataField pt.PivotFields("Units"), "Sum of Units", xlSum'),
}


def install_steps(table: dict[str, tuple[str, str, str]]) -> str:
    """The Scenarios macro running each of ``table``'s scenarios, W reading the widths again."""
    steps = []
    for index, (_rows, _columns, step) in enumerate(table.values(), 1):
        lines = [f"    Case {index}"]
        for line in step.splitlines():
            lines.append('        out = out & "w=" & Widths(ws) & " "' if line == "W" else "        " + line)
        steps.append("\n".join(lines) + "\n")
    calls = "".join(f'    out = out & Run("{name}", "{rows}", "{columns}", {index})\n'
                    for index, (name, (rows, columns, _step)) in enumerate(table.items(), 1))
    return SOURCE.replace("{steps}", "".join(steps)).replace("{calls}", calls)


def session() -> ExcelSession:
    return ExcelSession(HarnessConfig(lock_wait_s=1200.0))


def main() -> None:
    wanted = set(sys.argv[1:]) or {"styles", "compact", "subtotals", "layouts", "steps"}
    OUT.mkdir(parents=True, exist_ok=True)
    path = OUT / "pivot_column_fit.json"
    record = json.loads(path.read_text(encoding="utf-8")) if path.exists() else {}
    record.update({"texts": TEXTS, "grid": GRID, "source": SOURCE})
    if "styles" in wanted:
        with session() as excel:
            excel.new_document()
            result = excel.run_vba(STYLES, "Styles", timeout=300.0)
            assert result.ok, f"{result.outcome}: {result.message}"
        styles = {}
        for one in str(result.value).split("|"):
            if "~" in one:
                name, rest = one.split("~", 1)
                kinds, elements = rest.split(":", 1)
                styles[name] = {"kinds": kinds, "bold": [int(e) for e in elements.split(",") if e]}
        record["styles"] = styles
        print(len(styles), "styles")
    for part, totals in (("compact", False), ("subtotals", True)):
        if part not in wanted:
            continue
        found: dict[str, str] = {}
        with session() as excel:
            excel.new_document()
            for index, text in enumerate(TEXTS):
                result = excel.run_vba(GRID, "Grid", (index, text, totals), timeout=900.0)
                assert result.ok, f"{result.outcome}: {result.message}"
                found.update(one.split("=") for one in str(result.value).split(";") if "=" in one)
        record[part] = found
        print(part, len(found))
    for part, table in (("layouts", LAYOUTS), ("steps", STEPS)):
        if part not in wanted:
            continue
        with session() as excel:
            excel.new_document()
            result = excel.run_vba(install_steps(table), "Scenarios", timeout=900.0)
            assert result.ok, f"{result.outcome}: {result.message}"
        found = dict(one.split("^", 1) for one in str(result.value).split("|") if "^" in one)
        record[part] = {name: {"rows": rows, "columns": columns, "step": step, "widths": found.get(name, "")}
                        for name, (rows, columns, step) in table.items()}
        for name, entry in record[part].items():
            print(part, name, entry["widths"])
    path.write_text(json.dumps(record, indent=1, ensure_ascii=False) + "\n", encoding="utf-8")


if __name__ == "__main__":
    main()
