"""Sorting a pivot table's items, as Excel lays the table out, reads and saves it.

A macro sorts a field's items by PivotField.AutoSort, by their names or
by a data field's values, up or down or back to manual, and moves one by
PivotItem.Position. Every layout sits on a sheet of its own in one new
workbook, a pivot table of its own over one source list of regions,
products and amounts: rows by region, with products down the columns, on
the page or inside the regions, and the sum of the amounts. The dump
reads every cell of the table, each field's items in order with their
positions, and each field's AutoSortOrder and AutoSortField. A layout
that refreshes, or sorts values that tie, has a source of its own. The
workbook is saved, its pivot table parts kept, and Excel opens it and
dumps every table again.

    python scripts/measure_pivot_sorts.py

writes tests/fixtures/pivot_sorts/: the workbook as Excel saved it, and
pivot_sorts.json, which tests/test_excel_pivot_sorts.py replays.
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

OUT = ROOT / "tests" / "fixtures" / "pivot_sorts"

HELPER = '''Private note As String

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
                out = out & it.Name & "=" & it.Position & IIf(it.Visible, "", "h") & ","
            Next
            Err.Clear
            out = out & " order=" & f.AutoSortOrder
            If Err.Number <> 0 Then out = out & " order=E" & Err.Number
            Err.Clear
            out = out & " by=" & f.AutoSortField
            If Err.Number <> 0 Then out = out & " by=E" & Err.Number
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
    Dim ws As Object, pt As Object, out As String, it As Object, own As Object, held As Object
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = name
    Set own = cache
    If Len(source) > 0 Then
        Set held = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        held.Name = source
        held.Range("A1:C13").Value = wb.Worksheets("data").Range("A1:C13").Value
        If Left(source, 4) = "tied" Then held.Range("C3").Value = 50
        If Left(source, 6) = "months" Then held.Range("A2:A13").Value = Application.Transpose(Array("Mar", "Jan", _
            "Feb", "Apr", "Mar", "Jan", "Feb", "Apr", "Mar", "Jan", "Feb", "Apr"))
        Set own = wb.PivotCaches.Create(xlDatabase, source & "!R1C1:R13C3")
    End If
    Set pt = Made(wb, ws, own, rowField, columnField, pageField)
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


def probe(label: str, statement: str) -> str:
    """A statement, and its error number into the note under ``label`` if it raises one."""
    return (f"Err.Clear\n{statement}\n"
            f'If Err.Number <> 0 Then note = note & "{label}=E" & Err.Number & ";"\nErr.Clear\n')


#: name -> (the row field, the column field, the page field, a source sheet of the table's own or "" for the shared
#: one, what the layout does to its table).
LAYOUTS: dict[str, tuple[str, str, str, str, str]] = {
    "label_ascending": ("Region", "", "", "", f'{REGION}.AutoSort xlAscending, "Region"'),
    "label_descending": ("Region", "", "", "", f'{REGION}.AutoSort xlDescending, "Region"'),
    "value_ascending": ("Region", "", "", "", f'{REGION}.AutoSort xlAscending, "Sum of Amount"'),
    "value_descending": ("Region", "", "", "", f'{REGION}.AutoSort xlDescending, "Sum of Amount"'),
    "value_by_source": ("Region", "", "", "", probe("sort", f'{REGION}.AutoSort xlDescending, "Amount"')),
    "manual_after": ("Region", "", "", "", f'{REGION}.AutoSort xlDescending, "Sum of Amount"\n'
                                           f'{REGION}.AutoSort xlManual, "Region"'),
    "manual_by_value": ("Region", "", "", "", f'{REGION}.AutoSort xlDescending, "Region"\n'
                                              f'{REGION}.AutoSort xlManual, "Sum of Amount"'),
    "position_first": ("Region", "", "", "", f'{REGION}.PivotItems("West").Position = 1'),
    "position_last": ("Region", "", "", "", f'{REGION}.PivotItems("East").Position = 4'),
    "position_middle": ("Region", "", "", "", f'{REGION}.PivotItems("West").Position = 2'),
    "position_bad": ("Region", "", "", "", probe("zero", f'{REGION}.PivotItems("West").Position = 0') +
                     probe("five", f'{REGION}.PivotItems("West").Position = 5')),
    "position_after_sort": ("Region", "", "", "", f'{REGION}.AutoSort xlAscending, "Sum of Amount"\n'
                                                  f'{REGION}.PivotItems("South").Position = 1'),
    "sort_after_position": ("Region", "", "", "", f'{REGION}.PivotItems("West").Position = 1\n'
                                                  f'{REGION}.AutoSort xlAscending, "Region"'),
    "columns_by_value": ("Region", "Product", "", "", f'{REGION}.AutoSort xlDescending, "Sum of Amount"'),
    "column_labels": ("Region", "Product", "", "", f'{PRODUCT}.AutoSort xlDescending, "Product"'),
    "column_values": ("Region", "Product", "", "", f'{PRODUCT}.AutoSort xlDescending, "Sum of Amount"'),
    "inner_by_value": ("Region", "", "", "", INNER + f'{PRODUCT}.AutoSort xlDescending, "Sum of Amount"'),
    "outer_by_value": ("Region", "", "", "", INNER + f'{REGION}.AutoSort xlDescending, "Sum of Amount"'),
    "inner_labels": ("Region", "", "", "", INNER + f'{PRODUCT}.AutoSort xlDescending, "Product"'),
    "sort_then_hide": ("Region", "", "", "", f'{REGION}.AutoSort xlDescending, "Sum of Amount"\n'
                                             f'{REGION}.PivotItems("South").Visible = False'),
    "hide_then_sort": ("Region", "", "", "", f'{REGION}.PivotItems("South").Visible = False\n'
                                             f'{REGION}.AutoSort xlDescending, "Sum of Amount"'),
    "page_labels": ("Region", "", "Product", "", f'{PRODUCT}.AutoSort xlDescending, "Product"'),
    "page_values": ("Region", "", "Product", "", probe("sort", f'{PRODUCT}.AutoSort xlDescending, "Sum of Amount"')),
    "page_chosen_sorted": ("Region", "", "Product", "", f'{PRODUCT}.CurrentPage = "B"\n'
                                                        f'{REGION}.AutoSort xlDescending, "Sum of Amount"'),
    "hidden_field_sort": ("Region", "", "", "", probe("sort", f'{PRODUCT}.AutoSort xlDescending, "Product"') +
                          'For Each it In pt.PivotFields("Product").PivotItems: note = note & it.Name & "=" & '
                          'it.Position & ",": Next\n'
                          f'note = note & "order=" & {PRODUCT}.AutoSortOrder & ";"\n'
                          f'{PRODUCT}.Orientation = xlColumnField'),
    "bad_field": ("Region", "", "", "", probe("sort", f'{REGION}.AutoSort xlAscending, "Nope"')),
    "bad_order": ("Region", "", "", "", probe("sort", f'{REGION}.AutoSort 3, "Region"')),
    "other_row_field": ("Region", "", "", "", INNER + probe("sort", f'{REGION}.AutoSort xlAscending, "Product"')),
    "case_field": ("Region", "", "", "", probe("sort", f'{REGION}.AutoSort xlDescending, "sum of amount"')),
    "ties": ("Region", "", "", "tied", f'{REGION}.AutoSort xlAscending, "Sum of Amount"'),
    "ties_descending": ("Region", "", "", "tied2", f'{REGION}.AutoSort xlDescending, "Sum of Amount"'),
    "refresh_labels": ("Region", "", "", "data2", f'{REGION}.AutoSort xlAscending, "Region"\n'
                       'wb.Worksheets("data2").Range("A2").Value = "Central"\npt.PivotCache.Refresh'),
    "refresh_values": ("Region", "", "", "data3", f'{REGION}.AutoSort xlDescending, "Sum of Amount"\n'
                       'wb.Worksheets("data3").Range("C2").Value = 500\npt.PivotCache.Refresh'),
    "refresh_manual": ("Region", "", "", "data4", f'{REGION}.PivotItems("West").Position = 1\n'
                       'wb.Worksheets("data4").Range("A2").Value = "Central"\npt.PivotCache.Refresh'),
    "added_after_sort": ("Region", "", "", "", f'{REGION}.AutoSort xlDescending, "Sum of Amount"\n' + INNER),
    "position_hidden": ("Region", "", "", "", f'{REGION}.PivotItems("South").Visible = False\n'
                                              f'{REGION}.PivotItems("West").Position = 1'),
    "position_hidden_last": ("Region", "", "", "", f'{REGION}.PivotItems("South").Visible = False\n'
                                                   f'{REGION}.PivotItems("East").Position = 3'),
    "position_hidden_middle": ("Region", "", "", "", f'{REGION}.PivotItems("South").Visible = False\n'
                                                     f'{REGION}.PivotItems("East").Position = 2'),
    "position_hidden_up": ("Region", "", "", "", f'{REGION}.PivotItems("South").Visible = False\n'
                                                 f'{REGION}.PivotItems("West").Position = 2'),
    "position_hidden_bad": ("Region", "", "", "", f'{REGION}.PivotItems("South").Visible = False\n' +
                            probe("four", f'{REGION}.PivotItems("East").Position = 4') +
                            probe("hidden", f'{REGION}.PivotItems("South").Position = 1')),
    "position_type": ("Region", "", "", "", f'{REGION}.PivotItems("South").Visible = False\n'
                                            f'note = note & "type=" & VarType({REGION}.PivotItems("South").Position) & ";"\n'
                                            f'ws.Range("Z1").Value = {REGION}.PivotItems("South").Position\n'
                                            'note = note & "text=" & ws.Range("Z1").Text & ";"\nws.Range("Z1").Clear'),
    "two_data": ("Region", "", "", "tied3", 'pt.AddDataField pt.PivotFields("Amount"), "Max of Amount", xlMax\n'
                                            f'{REGION}.AutoSort xlAscending, "Max of Amount"'),
    "two_data_sum": ("Region", "", "", "tied4", 'pt.AddDataField pt.PivotFields("Amount"), "Max of Amount", xlMax\n'
                                                f'{REGION}.AutoSort xlAscending, "Sum of Amount"'),
    "months": ("Region", "", "", "monthsrc", f'{REGION}.AutoSort xlDescending, "Region"'),
    "months_initial": ("Region", "", "", "monthsrc2", ""),
}

BUILD = HELPER + '''
Public Function Build() As String
Dim wb As Object, data As Object, cache As Object, out As String, r As Long, regions As Variant, products As Variant
Dim ws As Object
Application.DisplayAlerts = False
Set wb = Workbooks.Add(xlWBATWorksheet)
{source}{layouts}
wb.Worksheets(1).Activate
wb.SaveAs "{folder}\\pivot_sorts.xlsx", 51
wb.Close False
Set wb = Workbooks.Open("{folder}\\pivot_sorts.xlsx")
out = out & "~"
For Each ws In wb.Worksheets
    If ws.PivotTables.Count > 0 Then out = out & ws.Name & "^" & Dump(ws.PivotTables(1)) & "|"
Next
wb.Close False
Build = out
End Function
'''

LAYOUT = 'out = out & Layout(wb, cache, "{name}", "{row}", "{column}", "{page}", "{source}", {step})\n'


def tables(path: Path) -> dict[str, tuple[str, str]]:
    """Each pivot table's pivotFields and rowItems and colItems, and its cache's cacheFields, as the file keeps them,
    by the table's name."""
    with zipfile.ZipFile(path) as package:
        out: dict[str, tuple[str, str]] = {}
        for part in sorted(name for name in package.namelist()
                           if re.fullmatch(r"xl/pivotTables/pivotTable\d+\.xml", name)):
            text = package.read(part).decode("utf-8")
            name = re.search(r'<pivotTableDefinition\b[^>]*?\bname="([^"]*)"', text)
            kept = "".join(found.group() for tag in ("pivotFields", "rowItems", "colItems")
                           if (found := re.search(rf"<{tag}\b.*?</{tag}>", text, re.DOTALL)))
            rels = package.read(part.replace("pivotTables/", "pivotTables/_rels/") + ".rels").decode("utf-8")
            target = re.search(r'Target="\.\./pivotCache/([^"]+)"', rels)
            cache = package.read(f"xl/pivotCache/{target.group(1)}").decode("utf-8") if target else ""
            cached = re.search(r"<cacheFields\b.*?</cacheFields>", cache, re.DOTALL)
            out[name.group(1) if name else part] = (kept, cached.group() if cached else "")
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
    strip_save_path(OUT / "pivot_sorts.xlsx")
    dumps = dict(one.split("^", 1) for one in session.split("|") if "^" in one)
    again = dict(one.split("^", 1) for one in reopened.split("|") if "^" in one)
    held = tables(OUT / "pivot_sorts.xlsx")
    record = {"build": BUILD, "source": SOURCE, "layout": LAYOUT, "steps": steps,
              "layouts": [{"name": name, "row": row, "column": column, "page": page, "own_source": source,
                           "step": step, "dump": dumps.get(name, ""), "reopened": again.get(name, ""),
                           "file": held.get(name, ("", ""))[0], "cache": held.get(name, ("", ""))[1]}
                          for name, (row, column, page, source, step) in LAYOUTS.items()]}
    (OUT / "pivot_sorts.json").write_text(json.dumps(record, indent=1, ensure_ascii=False) + "\n", encoding="utf-8")
    for layout in record["layouts"]:
        print("==", layout["name"])
        print("   session ", layout["dump"])
        if layout["reopened"] != layout["dump"]:
            print("   reopened", layout["reopened"])
        print("   file    ", re.sub(r"\s+", " ", str(layout["file"]))[:1200])


if __name__ == "__main__":
    main()
