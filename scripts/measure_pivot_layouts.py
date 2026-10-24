"""A pivot table's layout, compact, outline or tabular, and its fields' subtotals, as Excel lays the table out, reads
and saves it.

A macro lays a pivot table out by RowAxisLayout or by each field's
LayoutForm and LayoutCompactRow, and sets its fields' Subtotals,
LayoutSubtotalLocation, LayoutBlankLine and RepeatLabels, or the table's
RepeatAllLabels, header captions and grand total name. Every layout sits
on a sheet of its own in one new workbook, a pivot table of its own over
one source list of regions, products, channels, amounts and units, the
sum of the amounts on it. The dump reads every cell of the table, with
each cell's indent and whether it is merged, the table's ranges and
layout properties, and each field's: its orientation, position, layout,
subtotals and ranges. A layout that refreshes has a source of its own.
The workbook is saved, its pivot table parts kept, and Excel opens it and
dumps every table again.

    python scripts/measure_pivot_layouts.py

writes tests/fixtures/pivot_layouts/: the workbook as Excel saved it, and
pivot_layouts.json, which tests/test_excel_pivot_layouts.py replays.
"""

from __future__ import annotations

import json
import re
import shutil
import sys
import zipfile
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "scripts"))
from fixture_workbook import strip_save_path  # noqa: E402

OUT = ROOT / "tests" / "fixtures" / "pivot_layouts"

HELPER = '''Private note As String

Private Function Joined(v As Variant) As String
    Dim i As Long, out As String
    For i = LBound(v) To UBound(v)
        out = out & IIf(CBool(v(i)), "1", "0")
    Next
    Joined = LBound(v) & ":" & out
End Function

Private Function Props(o As Object, names As String) As String
    Dim n As Variant, out As String, v As Variant
    On Error Resume Next
    For Each n In Split(names, " ")
        Err.Clear
        v = Empty
        v = CallByName(o, CStr(n), VbGet)
        If Err.Number <> 0 Then
            out = out & " " & n & "=E" & Err.Number
        ElseIf IsArray(v) Then
            out = out & " " & n & "=" & Joined(v)
        Else
            out = out & " " & n & "=" & CStr(v)
        End If
    Next
    Props = out
End Function

Private Function Spots(o As Object, names As String) As String
    Dim n As Variant, out As String, r As Object
    On Error Resume Next
    For Each n In Split(names, " ")
        Err.Clear
        Set r = Nothing
        Set r = CallByName(o, CStr(n), VbGet)
        If Err.Number <> 0 Then
            out = out & " " & n & "=E" & Err.Number
        ElseIf r Is Nothing Then
            out = out & " " & n & "=-"
        Else
            out = out & " " & n & "=" & r.Address(False, False)
        End If
    Next
    Spots = out
End Function

Private Function Dump(pt As Object) As String
    Dim out As String, c As Object, f As Object
    On Error Resume Next
    out = "range=" & pt.TableRange2.Address(False, False) & " cells="
    For Each c In pt.TableRange2.Cells
        out = out & CStr(c.Value)
        If c.IndentLevel > 0 Then out = out & "~" & c.IndentLevel
        If c.MergeCells Then out = out & "~m"
        out = out & ","
    Next
    out = out & Spots(pt, "TableRange1 RowRange ColumnRange DataBodyRange DataLabelRange")
    out = out & Props(pt, "CompactLayoutRowHeader CompactLayoutColumnHeader GrandTotalName DisplayFieldCaptions " _
        & "CompactRowIndent MergeLabels LayoutRowDefault RowGrand ColumnGrand")
    For Each f In pt.PivotFields
        out = out & " " & f.Name & ":" & Props(f, "Orientation Position LayoutForm LayoutCompactRow " _
            & "LayoutSubtotalLocation LayoutBlankLine LayoutPageBreak RepeatLabels Subtotals SubtotalName")
        If f.Orientation = xlRowField Or f.Orientation = xlColumnField Then out = out & Spots(f, "LabelRange DataRange")
    Next
    Dump = out
End Function

Private Function Made(ws As Object, cache As Object, rowFields As String, columnFields As String, _
        pageFields As String) As Object
    Dim pt As Object, n As Variant
    Set pt = cache.CreatePivotTable(ws.Range("A3"), ws.Name)
    For Each n In Split(rowFields, ",")
        pt.PivotFields(CStr(n)).Orientation = xlRowField
    Next
    For Each n In Split(columnFields, ",")
        pt.PivotFields(CStr(n)).Orientation = xlColumnField
    Next
    For Each n In Split(pageFields, ",")
        pt.PivotFields(CStr(n)).Orientation = xlPageField
    Next
    pt.AddDataField pt.PivotFields("Amount"), "Sum of Amount", xlSum
    Set Made = pt
End Function

Private Function Layout(wb As Object, cache As Object, name As String, rowFields As String, columnFields As String, _
        pageFields As String, source As String, step As Long) As String
    Dim ws As Object, pt As Object, out As String, own As Object, held As Object, total As Object, x As Variant
    Dim r As Long
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = name
    Set own = cache
    If Len(source) > 0 Then
        Set held = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        held.Name = source
        held.Range("A1:E17").Value = wb.Worksheets("data").Range("A1:E17").Value
        If Left(source, 9) = "src_dates" Then
            For r = 2 To 17
                held.Cells(r, 3).Value = DateSerial(2025, 1 + (r Mod 3), 10 + (r Mod 2))
            Next
            held.Range("C2:C17").NumberFormat = "m/d/yyyy"
        End If
        If Left(source, 9) = "src_money" Then held.Range("D2:D17").NumberFormat = "$#,##0.00"
        Set own = wb.PivotCaches.Create(xlDatabase, source & "!R1C1:R17C5")
    End If
    Set pt = Made(ws, own, rowFields, columnFields, pageFields)
    Set total = pt.DataFields(1)
    note = ""
    On Error Resume Next
    Err.Clear
    Select Case step
{steps}
    End Select
    If Err.Number <> 0 Then out = "E" & Err.Number & " "
    If Len(note) > 0 Then out = out & "note=" & note & " "
    On Error GoTo 0
    Layout = name & "^" & out & Dump(pt) & "|"
End Function
'''

#: The source list: sixteen rows of a region, a product, a channel, an amount and units, West without cherries
#: and North without bananas.
SOURCE = '''Set data = wb.Worksheets(1)
data.Name = "data"
data.Range("A1:E1").Value = Array("Region", "Product", "Channel", "Amount", "Units")
rows = Array("East|Apple|Online|10|1", "East|Banana|Store|20|2", "East|Cherry|Online|30|3", _
    "West|Apple|Store|40|4", "West|Banana|Online|50|5", "North|Apple|Online|60|6", "North|Cherry|Store|70|7", _
    "South|Banana|Store|80|8", "South|Cherry|Online|90|9", "East|Apple|Store|15|1", "West|Apple|Online|25|2", _
    "North|Apple|Store|35|3", "South|Apple|Store|45|4", "West|Banana|Store|55|5", "South|Banana|Online|65|6", _
    "North|Cherry|Online|75|7")
For r = 0 To 15
    parts = Split(rows(r), "|")
    data.Cells(r + 2, 1).Value = parts(0)
    data.Cells(r + 2, 2).Value = parts(1)
    data.Cells(r + 2, 3).Value = parts(2)
    data.Cells(r + 2, 4).Value = CLng(parts(3))
    data.Cells(r + 2, 5).Value = CLng(parts(4))
Next
Set cache = wb.PivotCaches.Create(xlDatabase, "data!R1C1:R17C5")
'''


def field(name: str) -> str:
    return f'pt.PivotFields("{name}")'


REGION, PRODUCT, CHANNEL = field("Region"), field("Product"), field("Channel")
OUTLINE, TABULAR, COMPACT = "pt.RowAxisLayout xlOutlineRow", "pt.RowAxisLayout xlTabularRow", "pt.RowAxisLayout xlCompactRow"
SUM_COUNT = "Array(False, True, True, False, False, False, False, False, False, False, False, False)"
ALL = "Array(False, True, True, True, True, True, True, True, True, True, True, True)"
UNITS = 'pt.AddDataField pt.PivotFields("Units"), "Sum of Units", xlSum'


def probe(text: str, statement: str) -> str:
    """A statement, and its error number into the note under ``text`` if it raises one."""
    return (f"Err.Clear\n{statement}\n"
            f'If Err.Number <> 0 Then note = note & "{text}=E" & Err.Number & ";"\nErr.Clear\n')


def lines(*statements: str) -> str:
    return "\n".join(statements)


TWO = "Region,Product"
THREE = "Region,Product,Channel"

#: name -> (the row fields, the column fields, the page fields, a source sheet of the table's own or "" for the
#: shared one, what the layout does to its table).
LAYOUTS: dict[str, tuple[str, str, str, str, str]] = {
    "compact": (TWO, "", "", "", ""),
    "outline": (TWO, "", "", "", OUTLINE),
    "tabular": (TWO, "", "", "", TABULAR),
    "compact_again": (TWO, "", "", "", lines(TABULAR, COMPACT)),
    "outline_then_tabular": (TWO, "", "", "", lines(OUTLINE, TABULAR)),
    "compact_bottom": (TWO, "", "", "", f"{REGION}.LayoutSubtotalLocation = xlAtBottom"),
    "outline_bottom": (TWO, "", "", "", lines(OUTLINE, f"{REGION}.LayoutSubtotalLocation = xlAtBottom")),
    "tabular_top": (TWO, "", "", "", lines(TABULAR, f"{REGION}.LayoutSubtotalLocation = xlAtTop")),
    "bottom_then_tabular": (TWO, "", "", "", lines(f"{REGION}.LayoutSubtotalLocation = xlAtBottom", TABULAR)),
    "no_subtotals": (TWO, "", "", "", f"{REGION}.Subtotals(1) = False"),
    "outline_no_subtotals": (TWO, "", "", "", lines(OUTLINE, f"{REGION}.Subtotals(1) = False")),
    "tabular_no_subtotals": (TWO, "", "", "", lines(TABULAR, f"{REGION}.Subtotals(1) = False")),
    "sum_count": (TWO, "", "", "", f"{REGION}.Subtotals = {SUM_COUNT}"),
    "outline_sum_count": (TWO, "", "", "", lines(OUTLINE, f"{REGION}.Subtotals = {SUM_COUNT}")),
    "tabular_sum_count": (TWO, "", "", "", lines(TABULAR, f"{REGION}.Subtotals = {SUM_COUNT}")),
    "sum_count_top": (TWO, "", "", "", lines(f"{REGION}.Subtotals = {SUM_COUNT}",
                                             f"{REGION}.LayoutSubtotalLocation = xlAtTop")),
    "count_only": (TWO, "", "", "", f"{REGION}.Subtotals(3) = True"),
    "average_only": (TWO, "", "", "", f"{REGION}.Subtotals(4) = True"),
    "automatic_again": (TWO, "", "", "", lines(f"{REGION}.Subtotals(3) = True", f"{REGION}.Subtotals(1) = True")),
    "all_functions": (TWO, "", "", "", f"{REGION}.Subtotals = {ALL}"),
    "tabular_all_functions": (TWO, "", "", "", lines(TABULAR, f"{REGION}.Subtotals = {ALL}")),
    "automatic_and_sum": (TWO, "", "", "", probe("set", f"{REGION}.Subtotals = Array(True, True, False, False, "
                                                        "False, False, False, False, False, False, False, False)")),
    "short_array": (TWO, "", "", "", probe("set", f"{REGION}.Subtotals = Array(False, True)")),
    "none_false": (TWO, "", "", "", f"{REGION}.Subtotals = Array(False, False, False, False, False, False, "
                                    "False, False, False, False, False, False)"),
    "subtotal_bounds": (TWO, "", "", "", lines(probe("read0", f"x = {REGION}.Subtotals(0)"),
                                               probe("read13", f"x = {REGION}.Subtotals(13)"),
                                               probe("set13", f"{REGION}.Subtotals(13) = True"),
                                               f'note = note & "one=" & {REGION}.Subtotals(1) & ";"')),
    "data_subtotals": (TWO, "", "", "", lines(probe("read", "x = pt.DataFields(1).Subtotals(1)"),
                                              probe("set", "pt.DataFields(1).Subtotals(1) = False"))),
    "inner_no_subtotals": (TWO, "", "", "", f"{PRODUCT}.Subtotals(1) = False"),
    "inner_sum_count": (TWO, "", "", "", f"{PRODUCT}.Subtotals = {SUM_COUNT}"),
    "hidden_field_subtotals": (TWO, "", "", "", f"{CHANNEL}.Subtotals(1) = False"),
    "blank": (TWO, "", "", "", f"{REGION}.LayoutBlankLine = True"),
    "outline_blank": (TWO, "", "", "", lines(OUTLINE, f"{REGION}.LayoutBlankLine = True")),
    "tabular_blank": (TWO, "", "", "", lines(TABULAR, f"{REGION}.LayoutBlankLine = True")),
    "blank_inner": (TWO, "", "", "", f"{PRODUCT}.LayoutBlankLine = True"),
    "blank_no_subtotals": (TWO, "", "", "", lines(TABULAR, f"{REGION}.Subtotals(1) = False",
                                                  f"{REGION}.LayoutBlankLine = True")),
    "blank_bottom": (TWO, "", "", "", lines(f"{REGION}.LayoutSubtotalLocation = xlAtBottom",
                                            f"{REGION}.LayoutBlankLine = True")),
    "blank_single": ("Region", "", "", "", f"{REGION}.LayoutBlankLine = True"),
    "repeat_tabular": (TWO, "", "", "", lines(TABULAR, "pt.RepeatAllLabels xlRepeatLabels")),
    "repeat_outline": (TWO, "", "", "", lines(OUTLINE, "pt.RepeatAllLabels xlRepeatLabels")),
    "repeat_compact": (TWO, "", "", "", "pt.RepeatAllLabels xlRepeatLabels"),
    "repeat_off": (TWO, "", "", "", lines(TABULAR, "pt.RepeatAllLabels xlRepeatLabels",
                                          "pt.RepeatAllLabels xlDoNotRepeatLabels")),
    "repeat_field": (TWO, "", "", "", lines(TABULAR, f"{REGION}.RepeatLabels = True")),
    "repeat_three": (THREE, "", "", "", lines(TABULAR, "pt.RepeatAllLabels xlRepeatLabels")),
    "form_outer": (TWO, "", "", "", f"{REGION}.LayoutForm = xlTabular"),
    "form_inner": (TWO, "", "", "", f"{PRODUCT}.LayoutForm = xlTabular"),
    "compact_off_outer": (TWO, "", "", "", f"{REGION}.LayoutCompactRow = False"),
    "compact_off_inner": (TWO, "", "", "", f"{PRODUCT}.LayoutCompactRow = False"),
    "tabular_compact_on": (TWO, "", "", "", lines(TABULAR, f"{REGION}.LayoutCompactRow = True")),
    "outline_form_compact": (TWO, "", "", "", lines(TABULAR, f"{REGION}.LayoutForm = xlOutline",
                                                    f"{REGION}.LayoutCompactRow = True")),
    "three_compact": (THREE, "", "", "", ""),
    "three_outline": (THREE, "", "", "", OUTLINE),
    "three_tabular": (THREE, "", "", "", TABULAR),
    "three_outer_tabular": (THREE, "", "", "", f"{REGION}.LayoutForm = xlTabular"),
    "three_middle_tabular": (THREE, "", "", "", f"{PRODUCT}.LayoutForm = xlTabular"),
    "three_middle_outline": (THREE, "", "", "", f"{PRODUCT}.LayoutCompactRow = False"),
    "three_no_middle_subtotals": (THREE, "", "", "", lines(TABULAR, f"{PRODUCT}.Subtotals(1) = False")),
    "columns_compact": ("Region", "Product", "", "", ""),
    "columns_outline": ("Region", "Product", "", "", OUTLINE),
    "columns_tabular": ("Region", "Product", "", "", TABULAR),
    "two_columns": ("Region", "Product,Channel", "", "", ""),
    "two_columns_tabular": ("Region", "Product,Channel", "", "", TABULAR),
    "two_rows_two_columns_tabular": (TWO, "Channel", "", "", TABULAR),
    "column_no_subtotals": ("Region", "Product,Channel", "", "", f"{PRODUCT}.Subtotals(1) = False"),
    "column_sum_count": ("Region", "Product,Channel", "", "", f"{PRODUCT}.Subtotals = {SUM_COUNT}"),
    "column_blank": ("Region", "Product,Channel", "", "", f"{PRODUCT}.LayoutBlankLine = True"),
    "column_form": ("Region", "Product,Channel", "", "", f"{PRODUCT}.LayoutForm = xlTabular"),
    "values_rows": (TWO, "", "", "", lines(UNITS, "pt.DataPivotField.Orientation = xlRowField")),
    "values_rows_outline": (TWO, "", "", "", lines(UNITS, "pt.DataPivotField.Orientation = xlRowField", OUTLINE)),
    "values_rows_tabular": (TWO, "", "", "", lines(UNITS, "pt.DataPivotField.Orientation = xlRowField", TABULAR)),
    "values_columns_outline": (TWO, "", "", "", lines(UNITS, OUTLINE)),
    "values_columns_tabular": (TWO, "", "", "", lines(UNITS, TABULAR)),
    "values_field_layout": (TWO, "", "", "", lines(UNITS, probe("form", "pt.DataPivotField.LayoutForm = xlTabular"),
                                                   probe("subtotals", "x = pt.DataPivotField.Subtotals(1)"))),
    "single_outline": ("Region", "", "", "", OUTLINE),
    "single_tabular": ("Region", "", "", "", TABULAR),
    "column_only_tabular": ("", "Product", "", "", TABULAR),
    "headers": ("Region", "Product", "", "", lines('pt.CompactLayoutRowHeader = "Regions"',
                                                   'pt.CompactLayoutColumnHeader = "Products"')),
    "headers_tabular": ("Region", "Product", "", "", lines('pt.CompactLayoutRowHeader = "Regions"',
                                                           'pt.CompactLayoutColumnHeader = "Products"', TABULAR)),
    "grand_name": (TWO, "Channel", "", "", 'pt.GrandTotalName = "Everything"'),
    "captions_off": (TWO, "Channel", "", "", "pt.DisplayFieldCaptions = False"),
    "captions_off_tabular": (TWO, "Channel", "", "", lines(TABULAR, "pt.DisplayFieldCaptions = False")),
    "subtotal_name": (TWO, "", "", "", probe("name", f'{REGION}.SubtotalName = "Sub"')),
    "indent": (TWO, "", "", "", "pt.CompactRowIndent = 3"),
    "page_tabular": (TWO, "", "Channel", "", TABULAR),
    "hidden_tabular": (TWO, "", "", "", lines(TABULAR, f'{REGION}.PivotItems("East").Visible = False')),
    "sorted_tabular": (TWO, "", "", "", lines(TABULAR, f'{REGION}.AutoSort xlDescending, "Sum of Amount"')),
    "filtered_tabular": (TWO, "", "", "", lines(TABULAR, f"{REGION}.PivotFilters.Add Type:=xlValueIsGreaterThan, "
                                                         "DataField:=total, Value1:=150")),
    "default_then_add": ("Region", "", "", "", lines("pt.LayoutRowDefault = xlTabularRow",
                                                     f"{PRODUCT}.Orientation = xlRowField")),
    "tabular_then_add": ("Region", "", "", "", lines(TABULAR, f"{PRODUCT}.Orientation = xlRowField")),
    "tabular_then_move": (TWO, "", "", "", lines(TABULAR, f"{REGION}.Orientation = xlColumnField")),
    "refreshed_tabular": (TWO, "", "", "src_refresh", lines(
        TABULAR, f"{REGION}.Subtotals = {SUM_COUNT}",
        'wb.Worksheets("src_refresh").Range("A3").Value = "Central"', "pt.PivotCache.Refresh")),
    "layout_errors": (TWO, "", "", "", lines(probe("axis", "pt.RowAxisLayout 7"),
                                             probe("form", f"{REGION}.LayoutForm = 3"),
                                             probe("at", f"{REGION}.LayoutSubtotalLocation = 3"),
                                             probe("data", "pt.DataFields(1).LayoutForm = xlTabular"),
                                             probe("repeat", "pt.RepeatAllLabels 3"))),
    "subtotal_name_bottom": (TWO, "", "", "", lines(TABULAR, f'{REGION}.SubtotalName = "Sub"')),
    "subtotal_name_compact": (TWO, "", "", "", lines(f"{REGION}.LayoutSubtotalLocation = xlAtBottom",
                                                     f'{REGION}.SubtotalName = "Sub"')),
    "subtotal_name_columns": ("Region", "Product,Channel", "", "", f'{PRODUCT}.SubtotalName = "Sub"'),
    "sum_count_two_data": (TWO, "", "", "", lines(UNITS, f"{REGION}.Subtotals = {SUM_COUNT}")),
    "sum_count_values_rows": (TWO, "", "", "", lines(UNITS, "pt.DataPivotField.Orientation = xlRowField",
                                                     f"{REGION}.Subtotals = {SUM_COUNT}")),
    "count_data_sum_subtotal": (TWO, "", "", "", lines("total.Function = xlCount", f"{REGION}.Subtotals(2) = True")),
    "column_compact_header": ("Region", "Product", "", "", lines(TABULAR, f"{PRODUCT}.LayoutCompactRow = True")),
    "row_compact_header": ("Region", "Product", "", "", lines(TABULAR, f"{REGION}.LayoutCompactRow = True")),
    "tabular_no_data": (TWO, "", "", "", lines(TABULAR, "total.Orientation = xlHidden")),
    "tabular_no_grand": (TWO, "Channel", "", "", lines(TABULAR, "pt.RowGrand = False", "pt.ColumnGrand = False")),
    "outline_inner_bottom": (THREE, "", "", "", lines(OUTLINE, f"{PRODUCT}.LayoutSubtotalLocation = xlAtBottom")),
    "page_break": (TWO, "", "", "", f"{REGION}.LayoutPageBreak = True"),
    "dates_compact": ("Channel,Region", "", "", "src_dates1", ""),
    "dates_outline": ("Channel,Region", "", "", "src_dates2", OUTLINE),
    "dates_tabular": ("Channel,Region", "", "", "src_dates3", TABULAR),
    "dates_inner_tabular": ("Region,Channel", "", "", "src_dates4", TABULAR),
    "money_tabular": (TWO, "", "", "src_money1", TABULAR),
    "money_compact": (TWO, "", "", "src_money2", ""),
    "indent_three": (THREE, "", "", "", "pt.CompactRowIndent = 2"),
    "indent_errors": (TWO, "", "", "", lines(probe("zero", "pt.CompactRowIndent = 0"),
                                             probe("big", "pt.CompactRowIndent = 128"),
                                             probe("most", "pt.CompactRowIndent = 127"))),
    "inner_sum_count_tabular": (TWO, "", "", "", lines(TABULAR, f"{PRODUCT}.Subtotals = {SUM_COUNT}")),
    "inner_sum_count_three": (THREE, "", "", "", f"{CHANNEL}.Subtotals = {SUM_COUNT}"),
    "inner_count_only": (TWO, "", "", "", f"{PRODUCT}.Subtotals(3) = True"),
    "inner_sum_count_hidden": (TWO, "", "", "", lines(f"{PRODUCT}.Subtotals = {SUM_COUNT}",
                                                      f'{PRODUCT}.PivotItems("Banana").Visible = False')),
    "column_inner_sum_count": ("Region", "Product,Channel", "", "", f"{CHANNEL}.Subtotals = {SUM_COUNT}"),
    "single_sum_count": ("Region", "", "", "", f"{REGION}.Subtotals = {SUM_COUNT}"),
    "indent_zero": (TWO, "", "", "", "pt.CompactRowIndent = 0"),
    "subtotal_name_sum_count": (TWO, "", "", "", lines(TABULAR, f"{REGION}.Subtotals = {SUM_COUNT}",
                                                       f'{REGION}.SubtotalName = "Sub"')),
    "function_average": (TWO, "", "", "", "total.Function = xlAverage"),
    "function_errors": (TWO, "", "", "", lines(probe("field", f"{REGION}.Function = xlSum"),
                                               probe("unknown", "total.Function = 99"),
                                               probe("text", 'total.Function = "x"'))),
    "function_custom_name": (TWO, "", "", "", lines('pt.AddDataField pt.PivotFields("Units"), "Money", xlSum',
                                                    'pt.DataFields("Money").Function = xlMax')),
    "function_values_rows": (TWO, "", "", "", lines(UNITS, "pt.DataPivotField.Orientation = xlRowField",
                                                    "total.Function = xlMin")),
    "merge_labels": (TWO, "", "", "", lines(TABULAR, "pt.MergeLabels = True")),
}

BUILD = HELPER + '''
Public Function Build() As String
Dim wb As Object, data As Object, cache As Object, out As String, r As Long, rows As Variant, parts As Variant
Dim ws As Object
Application.DisplayAlerts = False
Set wb = Workbooks.Add(xlWBATWorksheet)
{source}{layouts}
wb.Worksheets(1).Activate
wb.SaveAs "{folder}\\pivot_layouts.xlsx", 51
wb.Close False
Set wb = Workbooks.Open("{folder}\\pivot_layouts.xlsx")
out = out & "@@"
For Each ws In wb.Worksheets
    If ws.PivotTables.Count > 0 Then out = out & ws.Name & "^" & Dump(ws.PivotTables(1)) & "|"
Next
wb.Close False
Build = out
End Function
'''

LAYOUT = 'out = out & Layout(wb, cache, "{name}", "{row}", "{column}", "{page}", "{source}", {step})\n'

#: The parts of a pivot table's definition its layout shows in: the root's attributes, the location, the fields,
#: lines and the table's own extensions.
KEPT = ("location", "pivotFields", "rowFields", "rowItems", "colFields", "colItems", "extLst")


def tables(path: Path) -> dict[str, str]:
    """Each pivot table's definition as far as its layout shows in it, by the table's name."""
    with zipfile.ZipFile(path) as package:
        out: dict[str, str] = {}
        for part in sorted(name for name in package.namelist()
                           if re.fullmatch(r"xl/pivotTables/pivotTable\d+\.xml", name)):
            text = package.read(part).decode("utf-8")
            name = re.search(r'<pivotTableDefinition\b[^>]*?\bname="([^"]*)"', text)
            start = re.search(r"<pivotTableDefinition\b[^>]*>", text)
            body = re.sub(r"<pivotFields\b.*?</pivotFields>", "", text, flags=re.DOTALL)
            kept = [start.group() if start else ""]
            for tag in KEPT:
                source = text if tag == "pivotFields" else body
                pattern = rf"<{tag}\b[^>]*/>|<{tag}\b.*?</{tag}>"
                if tag == "extLst":
                    pattern = r"<extLst>(?:(?!</extLst>).)*</extLst>(?=</pivotTableDefinition>)"
                found = re.search(pattern, source, re.DOTALL)
                if found:
                    kept.append(found.group())
            out[name.group(1) if name else part] = "".join(kept)
    return out


def main() -> None:
    if OUT.exists():
        shutil.rmtree(OUT)
    OUT.mkdir(parents=True)
    steps = "".join(f"    Case {index}\n" + "".join(f"        {line}\n" for line in step.splitlines())
                    for index, (_name, (_row, _column, _page, _source, step)) in enumerate(LAYOUTS.items(), 1))
    body = "".join(LAYOUT.replace("{name}", name).replace("{row}", row).replace("{column}", column)
                   .replace("{page}", page).replace("{source}", source).replace("{step}", str(index))
                   for index, (name, (row, column, page, source, _step)) in enumerate(LAYOUTS.items(), 1))
    code = (BUILD.replace("{steps}", steps).replace("{source}", SOURCE).replace("{layouts}", body)
            .replace("{folder}", str(OUT)))
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(code, "Build", timeout=600.0)
        assert result.ok, f"{result.outcome}: {result.message} {result.error}"
    session, reopened = str(result.value).split("@@", 1)
    strip_save_path(OUT / "pivot_layouts.xlsx")
    dumps = dict(one.split("^", 1) for one in session.split("|") if "^" in one)
    again = dict(one.split("^", 1) for one in reopened.split("|") if "^" in one)
    held = tables(OUT / "pivot_layouts.xlsx")
    record = {"build": BUILD, "source": SOURCE, "layout": LAYOUT, "steps": steps,
              "layouts": [{"name": name, "row": row, "column": column, "page": page, "own_source": source,
                           "step": step, "dump": dumps.get(name, ""), "reopened": again.get(name, ""),
                           "file": held.get(name, "")}
                          for name, (row, column, page, source, step) in LAYOUTS.items()]}
    (OUT / "pivot_layouts.json").write_text(json.dumps(record, indent=1, ensure_ascii=False) + "\n",
                                            encoding="utf-8")
    for layout in record["layouts"]:
        print("==", layout["name"])
        print("   session ", layout["dump"])
        if layout["reopened"] != layout["dump"]:
            print("   reopened", layout["reopened"])
        print("   file    ", re.sub(r"\s+", " ", str(layout["file"]))[:2500])


if __name__ == "__main__":
    main()
