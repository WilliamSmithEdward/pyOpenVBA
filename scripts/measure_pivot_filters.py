"""Hiding a pivot table's items and choosing its page field's item, as Excel lays the table out, reads and saves it.

A macro hides a pivot table's rows by PivotItem.Visible = False, and
filters it by its page field's CurrentPage, or, with
EnableMultiplePageItems, by hiding the page field's items too;
ClearAllFilters shows them all again. Every layout sits on a sheet of its
own in one new workbook, a pivot table of its own over one source list of
regions, products and amounts: rows by region and the sum of the amounts,
with products down the columns or on the page. The dump reads every cell
of the table, each field's items and whether each shows, the page
field's CurrentPage and the grand total. The workbook is saved, its pivot
table parts kept, and Excel opens it and dumps every table again.

    python scripts/measure_pivot_filters.py

writes tests/fixtures/pivot_filters/: the workbook as Excel saved it, and
pivot_filters.json, which tests/test_excel_pivot_filters.py replays.
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

OUT = ROOT / "tests" / "fixtures" / "pivot_filters"

HELPER = '''Private note As String

Private Function Dump(pt As Object) As String
    Dim out As String, c As Object, f As Object, it As Object, n As Long
    On Error Resume Next
    out = "range=" & pt.TableRange2.Address(False, False) & " cells="
    For Each c In pt.TableRange2.Cells
        out = out & CStr(c.Value) & ","
    Next
    For Each f In pt.PivotFields
        If f.Orientation <> xlHidden And f.Orientation <> xlDataField Then
            out = out & " " & f.Name & "@" & f.Orientation & ":"
            For Each it In f.PivotItems
                out = out & it.Name & "=" & it.Visible & ","
            Next
            Err.Clear
            n = f.VisibleItems.Count
            out = out & " shown=" & IIf(Err.Number <> 0, "E" & Err.Number, n)
            Err.Clear
            n = f.HiddenItems.Count
            out = out & " hidden=" & IIf(Err.Number <> 0, "E" & Err.Number, n)
            If f.Orientation = xlPageField Then
                Err.Clear
                out = out & " first=" & f.VisibleItems(1).Name
                If Err.Number <> 0 Then out = out & " first=E" & Err.Number
                Err.Clear
                out = out & " page=" & CStr(f.CurrentPage.Name)
                If Err.Number <> 0 Then out = out & " page=E" & Err.Number
                Err.Clear
                out = out & " multiple=" & f.EnableMultiplePageItems
            End If
        End If
    Next
    Err.Clear
    out = out & " total=" & CStr(pt.GetPivotData("Sum of Amount").Value)
    If Err.Number <> 0 Then out = out & " total=E" & Err.Number
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

EAST = 'pt.PivotFields("Region").PivotItems("East").Visible = False\n'
MULTIPLE = 'pt.PivotFields("Product").EnableMultiplePageItems = True\n'
HIDE_B = 'pt.PivotFields("Product").PivotItems("B").Visible = False\n'
#: Each of a field's items and whether it shows, into the note, for a field the dump leaves out.
ITEMS = ('For Each it In pt.PivotFields("{field}").PivotItems: note = note & it.Name & "=" & it.Visible & ",": Next\n'
         'note = note & ";"\n')


def probe(label: str, statement: str) -> str:
    """A statement, and its error number into the note under ``label`` if it raises one."""
    return (f"Err.Clear\n{statement}\n"
            f'If Err.Number <> 0 Then note = note & "{label}=E" & Err.Number & ";"\nErr.Clear\n')


#: name -> (the row field, the column field, the page field, a source sheet of the table's own or "" for the shared
#: one, what the layout does to its table).
LAYOUTS: dict[str, tuple[str, str, str, str, str]] = {
    "row_hidden": ("Region", "", "", "", EAST),
    "two_hidden": ("Region", "", "", "", EAST + 'pt.PivotFields("Region").PivotItems("West").Visible = False'),
    "shown_again": ("Region", "", "", "", EAST + 'pt.PivotFields("Region").PivotItems("East").Visible = True'),
    "all_hidden": ("Region", "", "", "", EAST + 'pt.PivotFields("Region").PivotItems("West").Visible = False\n'
                                         'pt.PivotFields("Region").PivotItems("North").Visible = False\n'
                                         'pt.PivotFields("Region").PivotItems("South").Visible = False'),
    "column_hidden": ("Region", "Product", "", "", 'pt.PivotFields("Product").PivotItems("C").Visible = False'),
    "page_item": ("Region", "", "Product", "", 'pt.PivotFields("Product").CurrentPage = "A"'),
    "page_all": ("Region", "", "Product", "", 'pt.PivotFields("Product").CurrentPage = "A"\n'
                                              'pt.PivotFields("Product").CurrentPage = "(All)"'),
    "page_multiple": ("Region", "", "Product", "", MULTIPLE + HIDE_B),
    "page_one_left": ("Region", "", "Product", "", MULTIPLE + HIDE_B +
                      'pt.PivotFields("Product").PivotItems("C").Visible = False'),
    "cleared": ("Region", "Product", "", "", EAST + 'pt.PivotFields("Product").PivotItems("C").Visible = False\n'
                                             'pt.ClearAllFilters'),
    "field_cleared": ("Region", "Product", "", "", EAST + 'pt.PivotFields("Product").PivotItems("C").Visible = False\n'
                                                   'pt.PivotFields("Region").ClearAllFilters'),
    "page_and_row": ("Region", "", "Product", "", 'pt.PivotFields("Product").CurrentPage = "B"\n'
                                                  'pt.PivotFields("Region").PivotItems("North").Visible = False'),
    "page_region": ("Product", "", "Region", "", 'pt.PivotFields("Region").CurrentPage = "West"'),
    "page_cleared": ("Region", "", "Product", "", 'pt.PivotFields("Product").CurrentPage = "C"\npt.ClearAllFilters'),
    "page_hidden_single": ("Region", "", "Product", "", HIDE_B),
    "page_shown_single": ("Region", "", "Product", "", 'pt.PivotFields("Product").CurrentPage = "A"\n'
                                                       'pt.PivotFields("Product").PivotItems("B").Visible = True'),
    "page_current_hidden": ("Region", "", "Product", "", 'pt.PivotFields("Product").CurrentPage = "A"\n'
                                                         'pt.PivotFields("Product").PivotItems("A").Visible = False'),
    "current_on_row": ("Region", "", "", "", probe("get", 'note = note & "get=" & '
                                                          'pt.PivotFields("Region").CurrentPage.Name & ";"') +
                       probe("set", 'pt.PivotFields("Region").CurrentPage = "East"') +
                       probe("multiple", 'note = note & "multiple=" & '
                                         'pt.PivotFields("Region").EnableMultiplePageItems & ";"')),
    "page_missing": ("Region", "", "Product", "", probe("set", 'pt.PivotFields("Product").CurrentPage = "Z"') +
                     probe("number", 'pt.PivotFields("Product").CurrentPage = 2')),
    "page_then_multiple": ("Region", "", "Product", "", 'pt.PivotFields("Product").CurrentPage = "A"\n' + MULTIPLE),
    "multiple_then_single": ("Region", "", "Product", "", MULTIPLE + HIDE_B +
                             'pt.PivotFields("Product").EnableMultiplePageItems = False'),
    "multiple_current": ("Region", "", "Product", "", MULTIPLE + HIDE_B +
                         'pt.PivotFields("Product").CurrentPage = "C"'),
    "multiple_cleared": ("Region", "", "Product", "", MULTIPLE + HIDE_B + 'pt.ClearAllFilters'),
    "page_field_cleared": ("Region", "", "Product", "", 'pt.PivotFields("Product").CurrentPage = "B"\n'
                                                        'pt.PivotFields("Product").ClearAllFilters'),
    "hidden_field_item": ("Region", "", "", "", 'pt.PivotFields("Product").PivotItems("C").Visible = False\n' +
                          ITEMS.replace("{field}", "Product")),
    "held_item": ("Region", "", "", "", 'Set it = pt.PivotFields("Region").PivotItems("East")\n' + EAST +
                  'note = note & "held=" & it.Visible & ";"'),
    "row_to_column": ("Region", "", "", "", EAST + 'pt.PivotFields("Region").Orientation = xlColumnField'),
    "row_removed_back": ("Region", "", "", "", EAST + 'pt.PivotFields("Region").Orientation = xlHidden\n' +
                         ITEMS.replace("{field}", "Region") + 'pt.PivotFields("Region").Orientation = xlRowField'),
    "row_to_page": ("Product", "Region", "", "", EAST + 'pt.PivotFields("Region").Orientation = xlPageField'),
    "page_to_row": ("Region", "", "Product", "", 'pt.PivotFields("Product").CurrentPage = "A"\n'
                                                 'pt.PivotFields("Product").Orientation = xlRowField'),
    "multiple_to_row": ("Region", "", "Product", "", MULTIPLE + HIDE_B +
                        'pt.PivotFields("Product").Orientation = xlRowField'),
    "inner_hidden": ("Region", "", "", "", 'pt.PivotFields("Product").Orientation = xlRowField\n'
                                           'pt.PivotFields("Product").PivotItems("C").Visible = False'),
    "refreshed": ("Region", "", "Product", "data2", EAST + 'pt.PivotFields("Product").CurrentPage = "B"\n'
                  'wb.Worksheets("data2").Range("A2:B2").Value = Array("Central", "B")\n'
                  'wb.Worksheets("data2").Range("B3").Value = "D"\npt.PivotCache.Refresh'),
    "refreshed_multiple": ("Region", "", "Product", "data3", MULTIPLE + HIDE_B +
                           'wb.Worksheets("data3").Range("B2").Value = "D"\npt.PivotCache.Refresh'),
    "refreshed_gone": ("Region", "", "", "data4", EAST + 'wb.Worksheets("data4").Range("A4").Value = "South"\n'
                       'wb.Worksheets("data4").Range("A8").Value = "South"\n'
                       'wb.Worksheets("data4").Range("A12").Value = "South"\npt.PivotCache.Refresh'),
    "refreshed_page_gone": ("Region", "", "Product", "data5", 'pt.PivotFields("Product").CurrentPage = "C"\n'
                            'wb.Worksheets("data5").Range("B4,B7,B10,B13").Value = "D"\npt.PivotCache.Refresh'),
    "refreshed_new_first": ("Region", "", "", "data6", 'wb.Worksheets("data6").Range("A3").Value = "Central"\n'
                            'pt.PivotCache.Refresh'),
    "refreshed_two_new": ("Region", "", "", "data7", 'wb.Worksheets("data7").Range("A2").Value = "Zed"\n'
                          'wb.Worksheets("data7").Range("A3").Value = "Alpha"\npt.PivotCache.Refresh'),
    "refreshed_two_gone": ("Region", "", "", "data8", 'wb.Worksheets("data8").Range("A3,A4,A7,A8,A11,A12").Value = '
                           '"South"\npt.PivotCache.Refresh'),
    "hidden_field_b": ("Region", "", "", "", 'pt.PivotFields("Product").PivotItems("B").Visible = False\n' +
                       ITEMS.replace("{field}", "Product")),
    "hidden_field_region": ("Product", "", "", "", 'pt.PivotFields("Region").PivotItems("West").Visible = False\n' +
                            ITEMS.replace("{field}", "Region")),
    "hidden_field_then_column": ("Region", "", "", "", 'pt.PivotFields("Product").PivotItems("C").Visible = False\n'
                                 'pt.PivotFields("Product").Orientation = xlColumnField'),
    "row_removed": ("Region", "Product", "", "", EAST + 'pt.PivotFields("Region").Orientation = xlHidden'),
    "page_current_middle": ("Region", "", "Product", "", 'pt.PivotFields("Product").CurrentPage = "B"\n'
                                                         'pt.PivotFields("Product").PivotItems("B").Visible = False'),
    "page_current_last": ("Region", "", "Product", "", 'pt.PivotFields("Product").CurrentPage = "C"\n'
                                                       'pt.PivotFields("Product").PivotItems("C").Visible = False'),
    "page_other_hidden": ("Region", "", "Product", "", 'pt.PivotFields("Product").CurrentPage = "A"\n'
                                                       'pt.PivotFields("Product").PivotItems("C").Visible = False'),
    "page_all_shown": ("Region", "", "Product", "", 'pt.PivotFields("Product").PivotItems("A").Visible = True'),
    "multiple_current_hidden": ("Region", "", "Product", "", MULTIPLE + HIDE_B +
                                'pt.PivotFields("Product").CurrentPage = "B"'),
    "multiple_all": ("Region", "", "Product", "", MULTIPLE + HIDE_B + 'pt.PivotFields("Product").CurrentPage = "(All)"'),
    "multiple_on_row": ("Region", "", "", "", probe("set", 'pt.PivotFields("Region").EnableMultiplePageItems = True') +
                        'note = note & "multiple=" & pt.PivotFields("Region").EnableMultiplePageItems & ";"'),
    "page_last_hidden": ("Region", "", "Product", "", MULTIPLE + 'pt.PivotFields("Product").PivotItems("A").Visible = False\n'
                         + HIDE_B + 'pt.PivotFields("Product").PivotItems("C").Visible = False'),
    "single_last_hidden": ("Region", "", "Product", "", 'pt.PivotFields("Product").PivotItems("A").Visible = False\n'
                           + HIDE_B + 'pt.PivotFields("Product").PivotItems("C").Visible = False'),
    "single_current_hidden": ("Region", "", "Product", "", HIDE_B +
                              probe("set", 'pt.PivotFields("Product").CurrentPage = "B"')),
    "hidden_field_names": ("Region", "", "", "", 'note = note & pt.PivotFields("Product").PivotItems("C").Name & ","'
                           ' & pt.PivotFields("Product").PivotItems(3).Name & ";"\n'
                           'pt.PivotFields("Product").PivotItems("C").Visible = False\n' +
                           ITEMS.replace("{field}", "Product") +
                           'pt.PivotFields("Product").PivotItems("B").Visible = False\n' +
                           ITEMS.replace("{field}", "Product") +
                           'pt.PivotFields("Product").PivotItems("A").Visible = True\n' +
                           ITEMS.replace("{field}", "Product")),
    "hidden_field_shown": ("Region", "", "", "", 'pt.PivotFields("Product").PivotItems("B").Visible = True\n' +
                           ITEMS.replace("{field}", "Product")),
    "lookup_twice": ("Region", "", "", "", 'note = note & pt.PivotFields("Product").PivotItems("C").Name & ","'
                     ' & pt.PivotFields("Product").PivotItems("C").Name & ";"'),
    "lookup_after_index": ("Region", "", "", "", 'note = note & pt.PivotFields("Product").PivotItems(3).Name & ","'
                           ' & pt.PivotFields("Product").PivotItems("C").Name & ";"'),
    "lookup_after_each": ("Region", "", "", "", ITEMS.replace("{field}", "Product") +
                          'note = note & pt.PivotFields("Product").PivotItems("C").Name & ";"'),
    "lookup_after_hide": ("Region", "", "", "", 'pt.PivotFields("Product").PivotItems(3).Visible = False\n'
                          'note = note & pt.PivotFields("Product").PivotItems("B").Name & ";"\n' +
                          ITEMS.replace("{field}", "Product")),
    "lookup_only": ("Region", "", "", "", 'note = note & pt.PivotFields("Product").PivotItems("C").Name & ";"'),
    "count_only": ("Region", "", "", "", 'note = note & pt.PivotFields("Product").PivotItems.Count & ";"'),
    "field_only": ("Region", "", "", "", 'note = note & pt.PivotFields("Product").Name & ";"'),
    "hidden_field_order": ("Product", "", "", "", ITEMS.replace("{field}", "Region") +
                           'note = note & pt.PivotFields("Region").PivotItems(2).Name & ";"'),
    "hidden_field_missing": ("Region", "", "", "", probe("missing", 'note = note & "missing=" & '
                                                                   'pt.PivotFields("Product").PivotItems("Z").Name & ";"') +
                             probe("listed", 'note = note & "listed=" & '
                                             'pt.PivotFields("Region").PivotItems("Z").Name & ";"') +
                             probe("number", 'note = note & "number=" & '
                                             'pt.PivotFields("Product").PivotItems(9).Name & ";"')),
    "page_all_object": ("Region", "", "Product", "", probe("value", 'note = note & "value=" & '
                                                                   'pt.PivotFields("Product").CurrentPage.Value & ";"') +
                        probe("visible", 'note = note & "visible=" & '
                                         'pt.PivotFields("Product").CurrentPage.Visible & ";"') +
                        probe("position", 'note = note & "position=" & '
                                          'pt.PivotFields("Product").CurrentPage.Position & ";"') +
                        probe("hidden", 'note = note & "hidden=" & '
                                        'pt.PivotFields("Product").HiddenItems(1).Name & ";"')),
}

#: Layouts whose sheet is deleted before the save: Excel saves a page field that allows several items and names one
#: (CurrentPage then EnableMultiplePageItems), and then cannot open the file it wrote.
UNSAVED = ("page_then_multiple",)

BUILD = HELPER + '''
Public Function Build() As String
Dim wb As Object, data As Object, cache As Object, out As String, r As Long, regions As Variant, products As Variant
Dim ws As Object
Application.DisplayAlerts = False
Set wb = Workbooks.Add(xlWBATWorksheet)
{source}{layouts}{unsaved}
wb.Worksheets(1).Activate
wb.SaveAs "{folder}\\pivot_filters.xlsx", 51
wb.Close False
Set wb = Workbooks.Open("{folder}\\pivot_filters.xlsx")
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
    """Each sheet's pivot table definition, its pivotFields and pageFields, and its cache's cacheFields, as the
    file keeps them."""
    with zipfile.ZipFile(path) as package:
        out: dict[str, tuple[str, str]] = {}
        for part in sorted(name for name in package.namelist() if re.fullmatch(r"xl/pivotTables/pivotTable\d+\.xml",
                                                                                 name)):
            text = package.read(part).decode("utf-8")
            name = re.search(r'<pivotTableDefinition\b[^>]*?\bname="([^"]*)"', text)
            fields = re.search(r"<pivotFields\b.*?</pivotFields>", text, re.DOTALL)
            pages = re.search(r"<pageFields\b.*?</pageFields>", text, re.DOTALL)
            rels = package.read(part.replace("pivotTables/", "pivotTables/_rels/") + ".rels").decode("utf-8")
            target = re.search(r'Target="\.\./pivotCache/([^"]+)"', rels)
            cache = package.read(f"xl/pivotCache/{target.group(1)}").decode("utf-8") if target else ""
            cached = re.search(r"<cacheFields\b.*?</cacheFields>", cache, re.DOTALL)
            out[name.group(1) if name else part] = ((fields.group() if fields else "") + (pages.group() if pages else ""),
                                                    cached.group() if cached else "")
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
    unsaved = "".join(f'wb.Worksheets("{name}").Delete\n' for name in UNSAVED)
    code = BUILD.replace("{steps}", steps).replace("{source}", SOURCE).replace("{layouts}", body).replace(
        "{unsaved}", unsaved).replace("{folder}", str(OUT))
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(code, "Build", timeout=600.0)
        assert result.ok, f"{result.outcome}: {result.message} {result.error}"
    session, reopened = str(result.value).split("~", 1)
    strip_save_path(OUT / "pivot_filters.xlsx")
    dumps = dict(one.split("^", 1) for one in session.split("|") if "^" in one)
    again = dict(one.split("^", 1) for one in reopened.split("|") if "^" in one)
    held = tables(OUT / "pivot_filters.xlsx")
    record = {"build": BUILD, "source": SOURCE, "layout": LAYOUT, "steps": steps, "unsaved": unsaved,
              "layouts": [{"name": name, "row": row, "column": column, "page": page, "own_source": source,
                           "step": step, "dump": dumps.get(name, ""), "reopened": again.get(name, ""),
                           "file": held.get(name, ("", ""))[0], "cache": held.get(name, ("", ""))[1]}
                          for name, (row, column, page, source, step) in LAYOUTS.items()]}
    (OUT / "pivot_filters.json").write_text(json.dumps(record, indent=1, ensure_ascii=False) + "\n", encoding="utf-8")
    for layout in record["layouts"]:
        print("==", layout["name"])
        print("   session ", layout["dump"])
        if layout["reopened"] != layout["dump"]:
            print("   reopened", layout["reopened"])
        print("   file    ", re.sub(r"\s+", " ", str(layout["file"]))[:900])
        if layout["own_source"]:
            print("   cache   ", re.sub(r"\s+", " ", str(layout["cache"]))[:900])


if __name__ == "__main__":
    main()
