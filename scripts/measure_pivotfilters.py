"""Filtering a pivot table's items by their names or values, as Excel lays the table out, reads and saves it.

A macro adds a label filter, a value filter or a top 10 filter to a
field's PivotFilters, clears them by ClearLabelFilters, ClearValueFilters
or ClearAllFilters, deletes one, or mixes one with hidden items. Every
layout sits on a sheet of its own in one new workbook, a pivot table of
its own over one source list of regions, products and amounts: rows by
region, with products down the columns, on the page or inside the
regions, and the sum of the amounts. The dump reads every cell of the
table, each field's items and whether each shows, and each field's
filters: their type, values, data field and whether active. A layout
that refreshes has a source of its own. The workbook is saved, its pivot
table parts kept, and Excel opens it and dumps every table again.

    python scripts/measure_pivotfilters.py

writes tests/fixtures/pivotfilters/: the workbook as Excel saved it, and
pivotfilters.json, which tests/test_excel_pivotfilters.py replays.
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

OUT = ROOT / "tests" / "fixtures" / "pivotfilters"

HELPER = '''Private note As String

Private Function Filters(f As Object) As String
    Dim out As String, pf As Object, n As Long
    On Error Resume Next
    Err.Clear
    n = f.PivotFilters.Count
    If Err.Number <> 0 Then Filters = " filters=E" & Err.Number: Exit Function
    out = " filters=" & n
    For Each pf In f.PivotFilters
        out = out & "[" & pf.FilterType
        Err.Clear
        out = out & " v1=" & CStr(pf.Value1)
        If Err.Number <> 0 Then out = out & " v1=E" & Err.Number
        Err.Clear
        out = out & " v2=" & CStr(pf.Value2)
        If Err.Number <> 0 Then out = out & " v2=E" & Err.Number
        Err.Clear
        out = out & " data=" & pf.DataField.Name
        If Err.Number <> 0 Then out = out & " data=E" & Err.Number
        Err.Clear
        out = out & " active=" & pf.Active
        If Err.Number <> 0 Then out = out & " active=E" & Err.Number
        Err.Clear
        out = out & " types=" & VarType(pf.Value1) & "/" & VarType(pf.Value2)
        If Err.Number <> 0 Then out = out & " types=E" & Err.Number
        out = out & "]"
    Next
    Filters = out
End Function

Private Function Dump(pt As Object) As String
    Dim out As String, c As Object, f As Object, it As Object
    On Error Resume Next
    out = "range=" & pt.TableRange2.Address(False, False) & " cells="
    For Each c In pt.TableRange2.Cells
        out = out & CStr(c.Value) & ","
    Next
    For Each f In pt.PivotFields
        If f.Orientation <> xlHidden And f.Orientation <> xlDataField Then
            out = out & " " & f.Name & "@" & f.Orientation & ":"
            For Each it In f.PivotItems
                out = out & it.Name & IIf(it.Visible, "", "-h") & ","
            Next
            out = out & Filters(f)
        End If
    Next
    Dump = out
End Function

Private Function Made(wb As Object, ws As Object, cache As Object, rowField As String, columnField As String, _
        pageField As String) As Object
    Dim pt As Object
    Set pt = cache.CreatePivotTable(ws.Range("A3"), ws.Name)
    pt.PivotFields(rowField).Orientation = xlRowField
    If Len(columnField) > 0 Then pt.PivotFields(columnField).Orientation = xlColumnField
    If Len(pageField) > 0 Then pt.PivotFields(pageField).Orientation = xlPageField
    pt.AddDataField pt.PivotFields("Amount"), "Sum of Amount", xlSum
    Set Made = pt
End Function

Private Function Layout(wb As Object, cache As Object, name As String, rowField As String, columnField As String, _
        pageField As String, source As String, step As Long) As String
    Dim ws As Object, pt As Object, out As String, it As Object, own As Object, held As Object, total As Object
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = name
    Set own = cache
    If Len(source) > 0 Then
        Set held = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        held.Name = source
        held.Range("A1:C13").Value = wb.Worksheets("data").Range("A1:C13").Value
        If Left(source, 4) = "tied" Then held.Range("C3").Value = 50
        Set own = wb.PivotCaches.Create(xlDatabase, source & "!R1C1:R13C3")
    End If
    Set pt = Made(wb, ws, own, rowField, columnField, pageField)
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

#: The source list: twelve rows of a region, a product and an amount.
SOURCE = '''Set data = wb.Worksheets(1)
data.Name = "data"
data.Range("A1:C1").Value = Array("Region", "Product", "Amount")
regions = Array("East", "West", "North", "South")
products = Array("A", "B", "C")
For r = 0 To 11
    data.Cells(r + 2, 1).Value = regions(r Mod 4)
    data.Cells(r + 2, 2).Value = products(r Mod 3)
    data.Cells(r + 2, 3).Value = (r + 1) * 10
Next
Set cache = wb.PivotCaches.Create(xlDatabase, "data!R1C1:R13C3")
'''

REGION = 'pt.PivotFields("Region")'
PRODUCT = 'pt.PivotFields("Product")'
INNER = 'pt.PivotFields("Product").Orientation = xlRowField\n'


def label(kind: str, first: str, second: str = "") -> str:
    """A label filter of ``kind`` on the regions."""
    return f'{REGION}.PivotFilters.Add Type:={kind}, Value1:="{first}"' + (f', Value2:="{second}"' if second else "")


def value(kind: str, first: object, second: object = None, field: str = REGION) -> str:
    """A value filter of ``kind`` on a field, by the sum of the amounts."""
    return (f'{field}.PivotFilters.Add Type:={kind}, DataField:=total, Value1:={first}'
            + (f", Value2:={second}" if second is not None else ""))


def probe(text: str, statement: str) -> str:
    """A statement, and its error number into the note under ``text`` if it raises one."""
    return (f"Err.Clear\n{statement}\n"
            f'If Err.Number <> 0 Then note = note & "{text}=E" & Err.Number & ";"\nErr.Clear\n')


#: name -> (the row field, the column field, the page field, a source sheet of the table's own or "" for the shared
#: one, what the layout does to its table).
LAYOUTS: dict[str, tuple[str, str, str, str, str]] = {
    "caption_equals": ("Region", "", "", "", label("xlCaptionEquals", "North")),
    "caption_not_equal": ("Region", "", "", "", label("xlCaptionDoesNotEqual", "North")),
    "begins_with": ("Region", "", "", "", label("xlCaptionBeginsWith", "N")),
    "not_begin_with": ("Region", "", "", "", label("xlCaptionDoesNotBeginWith", "N")),
    "ends_with": ("Region", "", "", "", label("xlCaptionEndsWith", "st")),
    "not_end_with": ("Region", "", "", "", label("xlCaptionDoesNotEndWith", "st")),
    "contains": ("Region", "", "", "", label("xlCaptionContains", "o")),
    "not_contain": ("Region", "", "", "", label("xlCaptionDoesNotContain", "o")),
    "contains_case": ("Region", "", "", "", label("xlCaptionContains", "O")),
    "caption_greater": ("Region", "", "", "", label("xlCaptionIsGreaterThan", "N")),
    "caption_greater_equal": ("Region", "", "", "", label("xlCaptionIsGreaterThanOrEqualTo", "North")),
    "caption_less": ("Region", "", "", "", label("xlCaptionIsLessThan", "South")),
    "caption_less_equal": ("Region", "", "", "", label("xlCaptionIsLessThanOrEqualTo", "North")),
    "caption_between": ("Region", "", "", "", label("xlCaptionIsBetween", "M", "T")),
    "caption_not_between": ("Region", "", "", "", label("xlCaptionIsNotBetween", "M", "T")),
    "equals_wildcard": ("Region", "", "", "", label("xlCaptionEquals", "*st")),
    "equals_question": ("Region", "", "", "", label("xlCaptionEquals", "?orth")),
    "contains_wildcard": ("Region", "", "", "", label("xlCaptionContains", "o*h")),
    "value_greater": ("Region", "", "", "", value("xlValueIsGreaterThan", 200)),
    "value_equals": ("Region", "", "", "", value("xlValueEquals", 180)),
    "value_not_equal": ("Region", "", "", "", value("xlValueDoesNotEqual", 180)),
    "value_greater_equal": ("Region", "", "", "", value("xlValueIsGreaterThanOrEqualTo", 210)),
    "value_less": ("Region", "", "", "", value("xlValueIsLessThan", 200)),
    "value_less_equal": ("Region", "", "", "", value("xlValueIsLessThanOrEqualTo", 180)),
    "value_between": ("Region", "", "", "", value("xlValueIsBetween", 160, 220)),
    "value_not_between": ("Region", "", "", "", value("xlValueIsNotBetween", 160, 220)),
    "top_count": ("Region", "", "", "", value("xlTopCount", 2)),
    "bottom_count": ("Region", "", "", "", value("xlBottomCount", 1)),
    "top_percent": ("Region", "", "", "", value("xlTopPercent", 50)),
    "bottom_percent": ("Region", "", "", "", value("xlBottomPercent", 20)),
    "top_sum": ("Region", "", "", "", value("xlTopSum", 300)),
    "bottom_sum": ("Region", "", "", "", value("xlBottomSum", 200)),
    "top_count_all": ("Region", "", "", "", value("xlTopCount", 9)),
    "top_count_zero": ("Region", "", "", "", probe("add", value("xlTopCount", 0))),
    "label_then_hidden": ("Region", "", "", "", label("xlCaptionContains", "o") + "\n"
                          f'{REGION}.PivotItems("North").Visible = False'),
    "hidden_then_label": ("Region", "", "", "", f'{REGION}.PivotItems("East").Visible = False\n'
                          + label("xlCaptionContains", "o")),
    "label_then_value": ("Region", "", "", "", label("xlCaptionContains", "o") + "\n"
                         + value("xlValueIsGreaterThan", 220)),
    "value_then_label": ("Region", "", "", "", value("xlValueIsGreaterThan", 160) + "\n"
                         + label("xlCaptionContains", "o")),
    "two_labels": ("Region", "", "", "", label("xlCaptionContains", "o") + "\n" + label("xlCaptionBeginsWith", "E")),
    "clear_label": ("Region", "", "", "", label("xlCaptionContains", "o") + f"\n{REGION}.ClearLabelFilters"),
    "clear_value_keeps_label": ("Region", "", "", "", label("xlCaptionContains", "o") + f"\n{REGION}.ClearValueFilters"),
    "clear_value": ("Region", "", "", "", value("xlValueIsGreaterThan", 200) + f"\n{REGION}.ClearValueFilters"),
    "clear_all": ("Region", "", "", "", value("xlValueIsGreaterThan", 200) + f"\n{REGION}.ClearAllFilters"),
    "table_clear_all": ("Region", "", "", "", label("xlCaptionContains", "o") + "\npt.ClearAllFilters"),
    "delete_filter": ("Region", "", "", "", label("xlCaptionContains", "o") + f"\n{REGION}.PivotFilters(1).Delete"),
    "inner_value": ("Region", "", "", "", INNER + value("xlValueIsGreaterThan", 80, field=PRODUCT)),
    "inner_top": ("Region", "", "", "", INNER + value("xlTopCount", 1, field=PRODUCT)),
    "outer_value_inner": ("Region", "", "", "", INNER + value("xlValueIsGreaterThan", 200)),
    "column_label": ("Region", "Product", "", "", f'{PRODUCT}.PivotFilters.Add Type:=xlCaptionEquals, Value1:="B"'),
    "column_value": ("Region", "Product", "", "", value("xlValueIsGreaterThan", 250, field=PRODUCT)),
    "row_value_with_columns": ("Region", "Product", "", "", value("xlTopCount", 2)),
    "page_label": ("Region", "", "Product", "", probe("add", f'{PRODUCT}.PivotFilters.Add Type:=xlCaptionEquals, '
                                                             'Value1:="B"')),
    "hidden_field_label": ("Region", "", "", "", probe("add", f'{PRODUCT}.PivotFilters.Add Type:=xlCaptionEquals, '
                                                              'Value1:="B"')),
    "bad_type": ("Region", "", "", "", probe("add", f'{REGION}.PivotFilters.Add Type:=99, Value1:="N"')),
    "value_no_data": ("Region", "", "", "", probe("add", f'{REGION}.PivotFilters.Add Type:=xlValueIsGreaterThan, '
                                                         'Value1:=200')),
    "label_number": ("Region", "", "", "", probe("add", f'{REGION}.PivotFilters.Add Type:=xlCaptionEquals, '
                                                        'Value1:=5')),
    "between_one_value": ("Region", "", "", "", probe("add", label("xlCaptionIsBetween", "M"))),
    "data_by_name": ("Region", "", "", "", f'{REGION}.PivotFilters.Add Type:=xlValueIsGreaterThan, '
                                           'DataField:=pt.PivotFields("Sum of Amount"), Value1:=200'),
    "filter_text_value": ("Region", "", "", "", f'{REGION}.PivotFilters.Add Type:=xlValueIsGreaterThan, '
                                                'DataField:=total, Value1:="200"'),
    "refreshed": ("Region", "", "", "data2", label("xlCaptionContains", "o") + "\n"
                  'wb.Worksheets("data2").Range("A2").Value = "Moon"\npt.PivotCache.Refresh'),
    "refreshed_value": ("Region", "", "", "data3", value("xlValueIsGreaterThan", 200) + "\n"
                        'wb.Worksheets("data3").Range("C2").Value = 500\npt.PivotCache.Refresh'),
    "sorted_then_top": ("Region", "", "", "", f'{REGION}.AutoSort xlDescending, "Region"\n'
                        + value("xlTopCount", 2)),
    "label_number_text": ("Region", "", "", "", label("xlCaptionEquals", "5")),
    "multiple_allowed": ("Region", "", "", "", "pt.AllowMultipleFilters = True\n" + label("xlCaptionContains", "o")
                         + "\n" + probe("second", label("xlCaptionBeginsWith", "N"))
                         + 'note = note & "allowed=" & pt.AllowMultipleFilters & ";"'),
    "clear_manual": ("Region", "", "", "", label("xlCaptionContains", "o") + "\n"
                     f'{REGION}.PivotItems("North").Visible = False\n{REGION}.ClearManualFilter'),
    "numeric_caption_greater": ("Amount", "", "", "", 'pt.PivotFields("Amount").PivotFilters.Add '
                                                      'Type:=xlCaptionIsGreaterThan, Value1:=100'),
    "numeric_caption_equals": ("Amount", "", "", "", 'pt.PivotFields("Amount").PivotFilters.Add '
                                                     'Type:=xlCaptionEquals, Value1:="20"'),
    "numeric_caption_contains": ("Amount", "", "", "", 'pt.PivotFields("Amount").PivotFilters.Add '
                                                       'Type:=xlCaptionContains, Value1:="1"'),
    "value_second_data": ("Region", "", "", "", 'pt.AddDataField pt.PivotFields("Amount"), "Max of Amount", xlMax\n'
                          f'{REGION}.PivotFilters.Add Type:=xlValueIsGreaterThan, DataField:=pt.DataFields(2), '
                          'Value1:=100'),
    "top_ties": ("Region", "", "", "tied", value("xlTopCount", 2)),
    "bottom_ties": ("Region", "", "", "tied2", value("xlBottomCount", 2)),
    "top_percent_exact": ("Region", "", "", "", value("xlTopPercent", "30.76923076923077")),
    "top_sum_exact": ("Region", "", "", "", value("xlTopSum", 240)),
}

BUILD = HELPER + '''
Public Function Build() As String
Dim wb As Object, data As Object, cache As Object, out As String, r As Long, regions As Variant, products As Variant
Dim ws As Object
Application.DisplayAlerts = False
Set wb = Workbooks.Add(xlWBATWorksheet)
{source}{layouts}
wb.Worksheets(1).Activate
wb.SaveAs "{folder}\\pivotfilters.xlsx", 51
wb.Close False
Set wb = Workbooks.Open("{folder}\\pivotfilters.xlsx")
out = out & "~"
For Each ws In wb.Worksheets
    If ws.PivotTables.Count > 0 Then out = out & ws.Name & "^" & Dump(ws.PivotTables(1)) & "|"
Next
wb.Close False
Build = out
End Function
'''

LAYOUT = 'out = out & Layout(wb, cache, "{name}", "{row}", "{column}", "{page}", "{source}", {step})\n'


def tables(path: Path) -> dict[str, str]:
    """Each pivot table's pivotFields, rowItems, colItems and filters as the file keeps them, by the table's name,
    with the table's own attributes."""
    with zipfile.ZipFile(path) as package:
        out: dict[str, str] = {}
        for part in sorted(name for name in package.namelist()
                           if re.fullmatch(r"xl/pivotTables/pivotTable\d+\.xml", name)):
            text = package.read(part).decode("utf-8")
            name = re.search(r'<pivotTableDefinition\b[^>]*?\bname="([^"]*)"', text)
            head = re.search(r"<pivotTableDefinition\b[^>]*>", text)
            kept = [found.group() for tag in ("pivotFields", "rowItems", "colItems")
                    if (found := re.search(rf"<{tag}\b.*?</{tag}>", text, re.DOTALL))]
            # The table's filters hold autoFilters with filters of their own; the outer list ends after a filter.
            filters = re.search(r'<filters count="\d+">.*?</filter></filters>', text, re.DOTALL)
            out[name.group(1) if name else part] = ((head.group() if head else "") + "".join(kept)
                                                    + (filters.group() if filters else ""))
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
    session, reopened = str(result.value).split("~", 1)
    strip_save_path(OUT / "pivotfilters.xlsx")
    dumps = dict(one.split("^", 1) for one in session.split("|") if "^" in one)
    again = dict(one.split("^", 1) for one in reopened.split("|") if "^" in one)
    held = tables(OUT / "pivotfilters.xlsx")
    record = {"build": BUILD, "source": SOURCE, "layout": LAYOUT, "steps": steps,
              "layouts": [{"name": name, "row": row, "column": column, "page": page, "own_source": source,
                           "step": step, "dump": dumps.get(name, ""), "reopened": again.get(name, ""),
                           "file": held.get(name, "")}
                          for name, (row, column, page, source, step) in LAYOUTS.items()]}
    (OUT / "pivotfilters.json").write_text(json.dumps(record, indent=1, ensure_ascii=False) + "\n", encoding="utf-8")
    for layout in record["layouts"]:
        print("==", layout["name"])
        print("   session ", layout["dump"])
        if layout["reopened"] != layout["dump"]:
            print("   reopened", layout["reopened"])
        print("   file    ", re.sub(r"\s+", " ", str(layout["file"]))[:1500])


if __name__ == "__main__":
    main()
