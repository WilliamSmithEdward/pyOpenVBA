"""A pivot table's data fields shown as Excel's calculations, Show Values As, as Excel works them out, reads and
saves them.

A macro sets a data field's Calculation, with its BaseField and BaseItem
where the calculation takes them: a difference from an item or from the
previous or next one, a percent of an item, of the row, column, total or a
parent, a running total, an index, or a rank. Every layout sits on a sheet
of its own in one new workbook, a pivot table of its own over one source
list of regions, products, channels, amounts and units: regions on the
rows, products on the columns, the sum of the amounts. The dump reads
every cell of the table with its number format, each data field's
calculation, base field, base item and number format. The workbook is
saved, its pivot table parts kept, and Excel opens it and dumps every
table again.

    python scripts/measure_pivot_calculations.py

writes tests/fixtures/pivot_calculations/: the workbook as Excel saved it,
and pivot_calculations.json, which tests/test_excel_pivot_calculations.py
replays.
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

OUT = ROOT / "tests" / "fixtures" / "pivot_calculations"

HELPER = '''Private note As String

Private Function Props(o As Object, names As String) As String
    Dim n As Variant, out As String, v As Variant
    On Error Resume Next
    For Each n In Split(names, " ")
        Err.Clear
        v = Empty
        v = CallByName(o, CStr(n), VbGet)
        If Err.Number <> 0 Then
            out = out & " " & n & "=E" & Err.Number
        Else
            out = out & " " & n & "=" & CStr(v)
        End If
    Next
    Props = out
End Function

Private Function Dump(pt As Object) As String
    Dim out As String, c As Object, f As Object
    On Error Resume Next
    out = "range=" & pt.TableRange2.Address(False, False) & " cells="
    For Each c In pt.TableRange2.Cells
        out = out & CStr(c.Value2)
        If VarType(c.Value) = vbError Then out = out & "~e"
        If c.NumberFormat <> "General" Then out = out & "~" & c.NumberFormat
        out = out & ","
    Next
    For Each f In pt.DataFields
        out = out & " " & f.Name & ":" & Props(f, "Calculation BaseField BaseItem NumberFormat Function")
    Next
    Dump = out
End Function

Private Function Labels(pt As Object) As String
    Dim c As Object, out As String
    On Error Resume Next
    For Each c In pt.RowRange.Cells
        out = out & CStr(c.Value) & "/"
    Next
    out = out & ":"
    For Each c In pt.ColumnRange.Cells
        out = out & CStr(c.Value) & "/"
    Next
    Labels = out
End Function

Private Function Made(ws As Object, cache As Object, rowFields As String, columnFields As String) As Object
    Dim pt As Object, n As Variant
    Set pt = cache.CreatePivotTable(ws.Range("A3"), ws.Name)
    For Each n In Split(rowFields, ",")
        pt.PivotFields(CStr(n)).Orientation = xlRowField
    Next
    For Each n In Split(columnFields, ",")
        pt.PivotFields(CStr(n)).Orientation = xlColumnField
    Next
    pt.AddDataField pt.PivotFields("Amount"), "Sum of Amount", xlSum
    Set Made = pt
End Function

Private Function Layout(wb As Object, cache As Object, name As String, rowFields As String, columnFields As String, _
        step As Long) As String
    Dim ws As Object, pt As Object, out As String, total As Object, x As Variant
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = name
    Set pt = Made(ws, cache, rowFields, columnFields)
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

#: The source list of tests/fixtures/pivot_layouts.
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


def probe(text: str, statement: str) -> str:
    return (f"Err.Clear\n{statement}\n"
            f'If Err.Number <> 0 Then note = note & "{text}=E" & Err.Number & ";"\nErr.Clear\n')


def lines(*statements: str) -> str:
    return "\n".join(statements)


def shown(calculation: str, field: str = "", item: str = "", *, steps: bool = False) -> str:
    """The data field shown as ``calculation``, then against a base field and item where given, each statement's
    error noted, and with ``steps`` the table's labels after each."""
    labels = '\nnote = note & "[" & Labels(pt) & "]"' if steps else ""
    out = [probe("calculation", f"total.Calculation = {calculation}") + labels]
    if field:
        out.append(probe("field", f'total.BaseField = "{field}"') + labels)
    if item:
        out.append(probe("item", f"total.BaseItem = {item}") + labels)
    return lines(*out)


def sorted_by(field: str, order: str = "xlDescending") -> str:
    """The field sorted by the data field's values, the table's labels noted after."""
    return f'pt.PivotFields("{field}").AutoSort {order}, "Sum of Amount"\nnote = note & "[" & Labels(pt) & "]"'


UNITS = 'pt.AddDataField pt.PivotFields("Units"), "Sum of Units", xlSum'
TWO = "Region,Product"
HIDE_BANANA = 'pt.PivotFields("Product").PivotItems("Banana").Visible = False'
BY_NAME = 'pt.PivotFields("Product").AutoSort xlDescending, "Product"'
COUNT_SUBTOTALS = 'pt.PivotFields("Region").Subtotals(3) = True'
NO_GRANDS = lines("pt.ColumnGrand = False", "pt.RowGrand = False")

#: name -> (the row fields, the column fields, what the layout does to its table).
LAYOUTS: dict[str, tuple[str, str, str]] = {
    "normal": ("Region", "Product", ""),
    "percent_of_total": ("Region", "Product", shown("xlPercentOfTotal")),
    "percent_of_row": ("Region", "Product", shown("xlPercentOfRow")),
    "percent_of_column": ("Region", "Product", shown("xlPercentOfColumn")),
    "index": ("Region", "Product", shown("xlIndex")),
    "percent_of_item": ("Region", "Product", shown("xlPercentOf", "Product", '"Apple"')),
    "percent_of_row_item": ("Region", "Product", shown("xlPercentOf", "Region", '"East"')),
    "difference_item": ("Region", "Product", shown("xlDifferenceFrom", "Product", '"Banana"')),
    "difference_previous": ("Region", "Product", shown("xlDifferenceFrom", "Product", '"(previous)"')),
    "difference_next": ("Region", "Product", shown("xlDifferenceFrom", "Product", '"(next)"')),
    "percent_difference": ("Region", "Product", shown("xlPercentDifferenceFrom", "Region", '"North"')),
    "percent_difference_previous": ("Region", "Product", shown("xlPercentDifferenceFrom", "Region",
                                                               '"(previous)"')),
    "running_total_rows": ("Region", "Product", shown("xlRunningTotal", "Region")),
    "running_total_columns": ("Region", "Product", shown("xlRunningTotal", "Product")),
    "percent_running_total": ("Region", "Product", shown("xlPercentRunningTotal", "Region")),
    "rank_ascending": ("Region", "Product", shown("xlRankAscending", "Region")),
    "rank_descending": ("Region", "Product", shown("xlRankDecending", "Product")),
    "parent_row": (TWO, "Channel", shown("xlPercentOfParentRow")),
    "parent_column": ("Region", "Product,Channel", shown("xlPercentOfParentColumn")),
    "parent_field": (TWO, "Channel", shown("xlPercentOfParent", "Region")),
    "two_rows_total": (TWO, "", shown("xlPercentOfTotal")),
    "two_rows_running": (TWO, "", shown("xlRunningTotal", "Product")),
    "two_rows_rank": (TWO, "", shown("xlRankDecending", "Product")),
    "two_rows_difference": (TWO, "", shown("xlDifferenceFrom", "Product", '"Apple"')),
    "two_data": ("Region", "Product", lines(UNITS, shown("xlPercentOfTotal"))),
    "back_to_normal": ("Region", "Product", lines(shown("xlPercentOfTotal"),
                                                  "total.Calculation = xlNoAdditionalCalculation")),
    "format_kept": ("Region", "Product", lines('total.NumberFormat = "0.0"', shown("xlPercentOfTotal"))),
    "rank_ties": ("Region", "Product", lines(UNITS, "Set total = pt.DataFields(2)",
                                             shown("xlRankDecending", "Product"))),
    "rank_ties_ascending": ("Region", "Product", lines(UNITS, "Set total = pt.DataFields(2)",
                                                       shown("xlRankAscending", "Product"))),
    "format_difference": ("Region", "Product", lines('total.NumberFormat = "0.0"',
                                                     shown("xlDifferenceFrom", "Product", '"Apple"'))),
    "percent_then_difference": ("Region", "Product", lines(shown("xlPercentOfTotal"),
                                                           shown("xlDifferenceFrom", "Product", '"Apple"'))),
    "format_index": ("Region", "Product", lines('total.NumberFormat = "0.0"', shown("xlIndex"))),
    "format_rank": ("Region", "Product", lines('total.NumberFormat = "0.0"', shown("xlRankAscending", "Region"))),
    "format_running": ("Region", "Product", lines('total.NumberFormat = "0.0"', shown("xlRunningTotal", "Region"))),
    "format_normal": ("Region", "Product", lines('total.NumberFormat = "0.0"',
                                                 shown("xlDifferenceFrom", "Product", '"Apple"'),
                                                 "total.Calculation = xlNoAdditionalCalculation")),
    "percent_then_running": ("Region", "Product", lines(shown("xlPercentOfTotal"),
                                                        shown("xlRunningTotal", "Region"))),
    "percent_then_index": ("Region", "Product", lines(shown("xlPercentOfTotal"), shown("xlIndex"))),
    "percent_then_rank": ("Region", "Product", lines(shown("xlPercentOfTotal"), shown("xlRankAscending", "Region"))),
    "own_percent_then_difference": ("Region", "Product", lines(shown("xlPercentOfTotal"), 'total.NumberFormat = "0%"',
                                                               shown("xlDifferenceFrom", "Product", '"Apple"'))),
    "own_format_then_percent_row": ("Region", "Product", lines(shown("xlPercentOfTotal"), 'total.NumberFormat = "0.0"',
                                                               shown("xlPercentOfRow"))),
    "two_rows_running_outer": (TWO, "", shown("xlRunningTotal", "Region")),
    "two_rows_parent_inner": (TWO, "", shown("xlPercentOfParent", "Product")),
    "two_rows_percent_outer": (TWO, "", shown("xlPercentOf", "Region", '"North"')),
    "two_rows_rank_outer": (TWO, "", shown("xlRankAscending", "Region")),
    "two_rows_percent_running_inner": (TWO, "", shown("xlPercentRunningTotal", "Product")),
    "tabular_difference": (TWO, "Channel", lines("pt.RowAxisLayout xlTabularRow",
                                                 shown("xlDifferenceFrom", "Channel", '"Online"'))),
    "base_off_axis": ("Region", "Product", shown("xlDifferenceFrom", "Channel", '"Online"')),
    "base_before": ("Region", "Product", lines(probe("field", 'total.BaseField = "Product"'),
                                               probe("item", 'total.BaseItem = "Apple"'))),
    "percent_of_previous": ("Region", "Product", shown("xlPercentOf", "Product", '"(previous)"')),
    "difference_columns_item": ("Region", "Product", shown("xlDifferenceFrom", "Region", '"West"')),
    "errors": ("Region", "Product", lines(probe("bad", "total.Calculation = 99"),
                                          probe("field", 'total.BaseField = "Nope"'),
                                          probe("item", 'total.BaseItem = "Nope"'),
                                          probe("onrow", 'pt.PivotFields("Region").Calculation = xlPercentOfTotal'),
                                          probe("difference_no_base", "total.Calculation = xlDifferenceFrom"))),
    # The previous or next item where a group lacks it, a hidden one, or another order.
    "two_rows_difference_previous": (TWO, "", shown("xlDifferenceFrom", "Product", '"(previous)"')),
    "two_rows_percent_next": (TWO, "", shown("xlPercentOf", "Product", '"(next)"')),
    "hidden_base_previous": ("Region", "Product", lines(HIDE_BANANA, shown("xlDifferenceFrom", "Product",
                                                                           '"(previous)"'))),
    "hidden_base_item": ("Region", "Product", lines(HIDE_BANANA, shown("xlDifferenceFrom", "Product", '"Banana"'))),
    "hidden_percent_total": ("Region", "Product", lines(HIDE_BANANA, shown("xlPercentOfTotal"))),
    "sorted_previous": ("Region", "Product", lines(BY_NAME, shown("xlDifferenceFrom", "Product", '"(previous)"'))),
    "sorted_base_item": ("Region", "Product", lines(BY_NAME, shown("xlDifferenceFrom", "Product", '"Apple"'))),
    "value_sorted_previous": ("Region", "Product", lines('pt.PivotFields("Product").AutoSort xlDescending, '
                                                         '"Sum of Amount"',
                                                         shown("xlDifferenceFrom", "Product", '"(previous)"'))),
    "value_sorted_running": ("Region", "Product", lines('pt.PivotFields("Region").AutoSort xlDescending, '
                                                        '"Sum of Amount"', shown("xlRunningTotal", "Region"))),
    # Cells empty on both sides, and the table's own empty and error text.
    "running_empty_start": (TWO, "Channel", shown("xlRunningTotal", "Channel")),
    "difference_both_empty": ("Region", "Product,Channel", shown("xlDifferenceFrom", "Region", '"North"')),
    "percent_both_empty": ("Region", "Product,Channel", shown("xlPercentOf", "Region", '"North"')),
    "percent_difference_both_empty": ("Region", "Product,Channel", shown("xlPercentDifferenceFrom", "Region",
                                                                         '"North"')),
    "null_string": ("Region", "Product", lines('pt.NullString = "-"', shown("xlPercentOf", "Product", '"Apple"'))),
    "error_string": ("Region", "Product", lines('pt.ErrorString = "x"', "pt.DisplayErrorString = True",
                                                shown("xlPercentOf", "Product", '"Apple"'))),
    # Subtotals by another function, grand totals hidden, the Values field on the rows, another function.
    "count_subtotals_total": (TWO, "", lines(COUNT_SUBTOTALS, shown("xlPercentOfTotal"))),
    "count_subtotals_running": (TWO, "", lines(COUNT_SUBTOTALS, shown("xlRunningTotal", "Region"))),
    "bottom_subtotals_parent": (TWO, "", lines('pt.PivotFields("Region").LayoutSubtotalLocation = xlAtBottom',
                                               shown("xlPercentOfParentRow"))),
    "no_grands_column": ("Region", "Product", lines(NO_GRANDS, shown("xlPercentOfColumn"))),
    "no_grands_index": ("Region", "Product", lines(NO_GRANDS, shown("xlIndex"))),
    "values_on_rows": ("Region", "Product", lines(UNITS, "pt.DataPivotField.Orientation = xlRowField",
                                                  "Set total = pt.DataFields(2)", shown("xlPercentOfColumn"))),
    "average_percent_total": ("Region", "Product", lines("total.Function = xlAverage", shown("xlPercentOfTotal"))),
    "base_page_field": ("Region", "Product", lines('pt.PivotFields("Channel").Orientation = xlPageField',
                                                   shown("xlDifferenceFrom", "Channel", '"Online"'))),
    # The base item a new base field keeps, and a base field given as a field.
    "base_item_kept": ("Region", "Product", lines(shown("xlDifferenceFrom", "Product", '"Cherry"'),
                                                  probe("again", 'total.BaseField = "Region"'))),
    "base_item_beyond": ("Region", "Channel", lines(shown("xlDifferenceFrom", "Product", '"Cherry"'),
                                                    probe("again", 'total.BaseField = "Channel"'))),
    "base_item_same_field": ("Region", "Product", lines(shown("xlDifferenceFrom", "Product", '"Cherry"'),
                                                        probe("again", 'total.BaseField = "Product"'))),
    "base_previous_new_field": ("Region", "Product", lines(shown("xlDifferenceFrom", "Product", '"(previous)"'),
                                                           probe("again", 'total.BaseField = "Region"'))),
    "two_rows_outer_previous": (TWO, "", shown("xlDifferenceFrom", "Region", '"(previous)"')),
    "two_rows_outer_next": (TWO, "", shown("xlPercentOf", "Region", '"(next)"')),
    # Leaving Excel 2010's calculations.
    "rank_then_normal": ("Region", "Product", lines(shown("xlRankAscending", "Region"),
                                                    "total.Calculation = xlNoAdditionalCalculation")),
    "rank_then_percent": ("Region", "Product", lines(shown("xlRankAscending", "Region"), shown("xlPercentOfTotal"))),
    "parent_then_difference": ("Region", "Product", lines(shown("xlPercentOfParentRow"),
                                                          shown("xlDifferenceFrom", "Product", '"Apple"'))),
    # A field sorted by the values of a data field shown as a calculation, the labels after each statement.
    "sort_steps_previous": ("Region", "Product", lines(sorted_by("Product"), shown("xlDifferenceFrom", "Product",
                                                                                     '"(previous)"', steps=True))),
    "sort_steps_running": ("Region", "Product", lines(sorted_by("Region"), shown("xlRunningTotal", "Region",
                                                                                   steps=True))),
    "sort_steps_running_ascending": ("Region", "Product", lines(sorted_by("Region", "xlAscending"),
                                                                shown("xlRunningTotal", "Region", steps=True))),
    "sort_percent_of_column": ("Region", "Product", lines(sorted_by("Product"), shown("xlPercentOfColumn",
                                                                                       steps=True))),
    "sort_index": ("Region", "Product", lines(sorted_by("Region"), shown("xlIndex", steps=True))),
    "sort_rank": ("Region", "Product", lines(sorted_by("Region"), shown("xlRankDecending", "Region", steps=True))),
    "sort_after_calculation": ("Region", "Product", lines(shown("xlPercentOfColumn", steps=True), sorted_by("Product"),
                                                          "pt.RowGrand = False", 'note = note & "[" & Labels(pt) & "]"',
                                                          "pt.RowGrand = True", 'note = note & "[" & Labels(pt) & "]"')),
    "sort_difference_item": ("Region", "Product", lines(sorted_by("Region"), shown("xlDifferenceFrom", "Region",
                                                                                     '"West"', steps=True))),
    "sort_inner_running": (TWO, "", lines(sorted_by("Product"), shown("xlRunningTotal", "Product", steps=True))),
    "previous_then_sort": ("Region", "Product", lines(shown("xlDifferenceFrom", "Product", '"(previous)"', steps=True),
                                                      sorted_by("Product"))),
    "next_sorted": ("Region", "Product", lines(sorted_by("Product"), shown("xlPercentOf", "Product", '"(next)"',
                                                                             steps=True))),
    "previous_rows_sorted": ("Region", "Product", lines(sorted_by("Region"), shown("xlPercentDifferenceFrom", "Region",
                                                                                     '"(previous)"', steps=True))),
    "previous_then_fixed": ("Region", "Product", lines(sorted_by("Product"), shown("xlDifferenceFrom", "Product",
                                                                                     '"(previous)"'),
                                                       probe("fixed", 'total.BaseItem = "Apple"'),
                                                       'note = note & "[" & Labels(pt) & "]"')),
    "sort_units_previous": ("Region", "Product", lines(UNITS, 'pt.PivotFields("Product").AutoSort xlDescending, '
                                                              '"Sum of Units"', 'note = note & "[" & Labels(pt) & "]"',
                                                       "Set total = pt.DataFields(2)",
                                                       shown("xlDifferenceFrom", "Product", '"(previous)"',
                                                             steps=True))),
    "sort_other_previous": ("Region", "Product", lines(UNITS, sorted_by("Product"), "Set total = pt.DataFields(2)",
                                                       shown("xlDifferenceFrom", "Product", '"(previous)"',
                                                             steps=True))),
    # A base field on no axis given without an item.
    "running_off_axis": ("Region", "Product", shown("xlRunningTotal", "Channel")),
    "difference_off_axis_no_item": ("Region", "Product", shown("xlDifferenceFrom", "Channel")),
    # Subtotals by another function under each calculation.
    "count_subtotals_row": (TWO, "Channel", lines(COUNT_SUBTOTALS, shown("xlPercentOfRow"))),
    "count_subtotals_column": (TWO, "Channel", lines(COUNT_SUBTOTALS, shown("xlPercentOfColumn"))),
    "count_subtotals_index": (TWO, "Channel", lines(COUNT_SUBTOTALS, shown("xlIndex"))),
    "count_subtotals_parent_row": (TWO, "Channel", lines(COUNT_SUBTOTALS, shown("xlPercentOfParentRow"))),
    "count_subtotals_parent_column": (TWO, "Channel", lines(COUNT_SUBTOTALS, shown("xlPercentOfParentColumn"))),
    "count_subtotals_parent": (TWO, "", lines(COUNT_SUBTOTALS, shown("xlPercentOfParent", "Region"))),
    "count_subtotals_difference": (TWO, "", lines(COUNT_SUBTOTALS, shown("xlDifferenceFrom", "Region", '"East"'))),
    "count_subtotals_previous": (TWO, "", lines(COUNT_SUBTOTALS, shown("xlPercentOf", "Region", '"(previous)"'))),
    "count_subtotals_rank": (TWO, "", lines(COUNT_SUBTOTALS, shown("xlRankDecending", "Region"))),
    "count_subtotals_percent_running": (TWO, "", lines(COUNT_SUBTOTALS, shown("xlPercentRunningTotal", "Region"))),
    "count_subtotals_inner": (TWO, "", lines(COUNT_SUBTOTALS, shown("xlRunningTotal", "Product"))),
    "count_bottom_parent_row": (TWO, "", lines(COUNT_SUBTOTALS,
                                               'pt.PivotFields("Region").LayoutSubtotalLocation = xlAtBottom',
                                               shown("xlPercentOfParentRow"))),
    "sum_count_subtotals_parent_row": (TWO, "", lines(
        'pt.PivotFields("Region").Subtotals = Array(False, True, True, False, False, False, False, False, False, '
        'False, False, False)', shown("xlPercentOfParentRow"))),
    "sum_count_subtotals_running": (TWO, "", lines(
        'pt.PivotFields("Region").Subtotals = Array(False, True, True, False, False, False, False, False, False, '
        'False, False, False)', shown("xlRunningTotal", "Region"))),
    "column_count_subtotals": ("Region", "Product,Channel", lines('pt.PivotFields("Product").Subtotals(3) = True',
                                                                  shown("xlPercentOfRow"))),
    "column_count_parent": ("Region", "Product,Channel", lines('pt.PivotFields("Product").Subtotals(3) = True',
                                                               shown("xlPercentOfParentColumn"))),
    "count_data_inner_sum": (TWO, "", lines("total.Function = xlCount",
                                            'pt.PivotFields("Region").Subtotals(2) = True',
                                            shown("xlPercentOfParentRow"))),
}

BUILD = HELPER + '''
Public Function Build() As String
Dim wb As Object, data As Object, cache As Object, out As String, r As Long, rows As Variant, parts As Variant
Dim ws As Object
Application.DisplayAlerts = False
Set wb = Workbooks.Add(xlWBATWorksheet)
{source}{layouts}
wb.Worksheets(1).Activate
wb.SaveAs "{folder}\\pivot_calculations.xlsx", 51
wb.Close False
Set wb = Workbooks.Open("{folder}\\pivot_calculations.xlsx")
out = out & "@@"
For Each ws In wb.Worksheets
    If ws.PivotTables.Count > 0 Then out = out & ws.Name & "^" & Dump(ws.PivotTables(1)) & "|"
Next
wb.Close False
Build = out
End Function
'''

LAYOUT = 'out = out & Layout(wb, cache, "{name}", "{row}", "{column}", {step})\n'


def tables(path: Path) -> dict[str, str]:
    """Each pivot table's data fields and extensions as the file keeps them, by the table's name."""
    with zipfile.ZipFile(path) as package:
        out: dict[str, str] = {}
        for part in sorted(name for name in package.namelist()
                           if re.fullmatch(r"xl/pivotTables/pivotTable\d+\.xml", name)):
            text = package.read(part).decode("utf-8")
            name = re.search(r'<pivotTableDefinition\b[^>]*?\bname="([^"]*)"', text)
            kept = [found.group() for pattern in (r"<dataFields\b.*?</dataFields>",
                                                  r"<extLst>(?:(?!</extLst>).)*</extLst>(?=</pivotTableDefinition>)")
                    if (found := re.search(pattern, text, re.DOTALL))]
            out[name.group(1) if name else part] = "".join(kept)
    return out


def main() -> None:
    if OUT.exists():
        shutil.rmtree(OUT)
    OUT.mkdir(parents=True)
    steps = "".join(f"    Case {index}\n" + "".join(f"        {line}\n" for line in step.splitlines())
                    for index, (_name, (_row, _column, step)) in enumerate(LAYOUTS.items(), 1))
    body = "".join(LAYOUT.replace("{name}", name).replace("{row}", row).replace("{column}", column)
                   .replace("{step}", str(index))
                   for index, (name, (row, column, _step)) in enumerate(LAYOUTS.items(), 1))
    code = (BUILD.replace("{steps}", steps).replace("{source}", SOURCE).replace("{layouts}", body)
            .replace("{folder}", str(OUT)))
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(code, "Build", timeout=600.0)
        assert result.ok, f"{result.outcome}: {result.message} {result.error}"
    session, reopened = str(result.value).split("@@", 1)
    strip_save_path(OUT / "pivot_calculations.xlsx")
    dumps = dict(one.split("^", 1) for one in session.split("|") if "^" in one)
    again = dict(one.split("^", 1) for one in reopened.split("|") if "^" in one)
    held = tables(OUT / "pivot_calculations.xlsx")
    record = {"build": BUILD, "source": SOURCE, "layout": LAYOUT, "steps": steps,
              "layouts": [{"name": name, "row": row, "column": column, "step": step,
                           "dump": dumps.get(name, ""), "reopened": again.get(name, ""), "file": held.get(name, "")}
                          for name, (row, column, step) in LAYOUTS.items()]}
    (OUT / "pivot_calculations.json").write_text(json.dumps(record, indent=1, ensure_ascii=False) + "\n",
                                                 encoding="utf-8")
    for layout in record["layouts"]:
        print("==", layout["name"])
        print("   session ", layout["dump"])
        if layout["reopened"] != layout["dump"]:
            print("   reopened", layout["reopened"])
        print("   file    ", re.sub(r"\s+", " ", str(layout["file"]))[:1500])


if __name__ == "__main__":
    main()
