"""A pivot table's items collapsed and expanded, ShowDetail, as Excel lays them out, reads and saves them.

A macro sets ShowDetail on an item or on a whole field, on the rows or the
columns, in compact, outline and tabular form, with subtotals at the top,
at the bottom or none. Every layout sits on a sheet of its own in one new
workbook, a pivot table of its own over one source list of regions,
products, channels, amounts and units: regions and products on the rows,
the sum of the amounts. The dump reads every cell of the table and the
ShowDetail of each field and each of its items. The workbook is saved, its
pivot table parts kept, and Excel opens it and dumps every table again.

    python scripts/measure_pivot_show_detail.py

writes tests/fixtures/pivot_show_detail/: the workbook as Excel saved it,
and pivot_show_detail.json, which tests/test_excel_pivot_show_detail.py
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

OUT = ROOT / "tests" / "fixtures" / "pivot_show_detail"

HELPER = '''Private note As String

Private Function Details(f As Object) As String
    Dim out As String, it As Object
    On Error Resume Next
    Err.Clear
    out = " [" & f.Name & " " & CStr(f.ShowDetail)
    If Err.Number <> 0 Then out = " [" & f.Name & " E" & Err.Number
    Err.Clear
    For Each it In f.PivotItems
        out = out & " " & it.Name & "=" & CStr(it.ShowDetail)
        If Err.Number <> 0 Then out = out & "E" & Err.Number
        Err.Clear
    Next
    Details = out & "]"
End Function

Private Function Dump(pt As Object) As String
    Dim out As String, c As Object, f As Object
    On Error Resume Next
    out = "range=" & pt.TableRange2.Address(False, False) & " cells="
    For Each c In pt.TableRange2.Cells
        out = out & CStr(c.Value2)
        If c.IndentLevel > 0 Then out = out & "~" & c.IndentLevel
        out = out & ","
    Next
    For Each f In pt.RowFields
        out = out & Details(f)
    Next
    For Each f In pt.ColumnFields
        If f.Name <> "Values" Then out = out & Details(f)
    Next
    Dump = out
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
    Dim ws As Object, pt As Object, out As String
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = name
    Set pt = Made(ws, cache, rowFields, columnFields)
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

#: The source list of tests/fixtures/pivot_calculations.
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


TWO = "Region,Product"
THREE = "Region,Product,Channel"
UNITS = 'pt.AddDataField pt.PivotFields("Units"), "Sum of Units", xlSum'
EAST = probe("east", 'pt.PivotFields("Region").PivotItems("East").ShowDetail = False')
REGION = probe("region", 'pt.PivotFields("Region").ShowDetail = False')

#: name -> (the row fields, the column fields, what the layout does to its table).
LAYOUTS: dict[str, tuple[str, str, str]] = {
    "normal": (TWO, "", ""),
    "collapse_item": (TWO, "", EAST),
    "collapse_field": (TWO, "", REGION),
    "collapse_then_expand": (TWO, "", lines(REGION, probe("north", 'pt.PivotFields("Region").PivotItems("North")'
                                                                    '.ShowDetail = True'))),
    "expand_field": (TWO, "", lines(REGION, probe("expand", 'pt.PivotFields("Region").ShowDetail = True'))),
    "outline_collapse": (TWO, "", lines("pt.RowAxisLayout xlOutlineRow", EAST)),
    "tabular_collapse": (TWO, "", lines("pt.RowAxisLayout xlTabularRow", EAST)),
    "tabular_collapse_field": (TWO, "", lines("pt.RowAxisLayout xlTabularRow", REGION)),
    "bottom_collapse": (TWO, "", lines('pt.PivotFields("Region").LayoutSubtotalLocation = xlAtBottom', EAST)),
    "no_subtotals_collapse": (TWO, "", lines('pt.PivotFields("Region").Subtotals(1) = False', EAST)),
    "inner_collapse": (TWO, "", probe("apple", 'pt.PivotFields("Product").PivotItems("Apple").ShowDetail = False')),
    "inner_field_collapse": (TWO, "", probe("product", 'pt.PivotFields("Product").ShowDetail = False')),
    "three_levels": (THREE, "", lines(EAST, probe("apple", 'pt.PivotFields("Product").PivotItems("Apple")'
                                                            '.ShowDetail = False'))),
    "three_middle_field": (THREE, "", probe("product", 'pt.PivotFields("Product").ShowDetail = False')),
    "columns_collapse": ("Region", "Product,Channel", probe("apple", 'pt.PivotFields("Product").PivotItems("Apple")'
                                                                    '.ShowDetail = False')),
    "columns_collapse_field": ("Region", "Product,Channel", probe("product", 'pt.PivotFields("Product").ShowDetail '
                                                                            '= False')),
    "hidden_after": (TWO, "", lines(EAST, 'pt.PivotFields("Region").PivotItems("North").Visible = False')),
    "hide_collapsed": (TWO, "", lines(EAST, 'pt.PivotFields("Region").PivotItems("East").Visible = False',
                                      'pt.PivotFields("Region").PivotItems("East").Visible = True')),
    "values_rows": (TWO, "", lines(UNITS, "pt.DataPivotField.Orientation = xlRowField", EAST)),
    "values_columns": (TWO, "", lines(UNITS, EAST)),
    "single_field": ("Region", "", probe("east", 'pt.PivotFields("Region").PivotItems("East").ShowDetail = False')),
    "errors": (TWO, "", lines(probe("data", 'pt.DataFields(1).ShowDetail = False'),
                              probe("values", 'pt.DataPivotField.ShowDetail = False'),
                              'pt.PivotFields("Channel").Orientation = xlPageField',
                              probe("page_item", 'pt.PivotFields("Channel").PivotItems("Online").ShowDetail = False'),
                              probe("page_field", 'pt.PivotFields("Channel").ShowDetail = False'),
                              probe("hidden_item", 'pt.PivotFields("Units").PivotItems("1").ShowDetail = False'),
                              probe("text", 'pt.PivotFields("Region").PivotItems("East").ShowDetail = "no"'))),
    "collapse_then_move": (TWO, "", lines(EAST, 'pt.PivotFields("Region").Orientation = xlColumnField')),
    "collapse_then_sort": (TWO, "", lines(EAST, 'pt.PivotFields("Region").AutoSort xlDescending, "Region"')),
}

BUILD = HELPER + '''
Public Function Build() As String
Dim wb As Object, data As Object, cache As Object, out As String, r As Long, rows As Variant, parts As Variant
Dim ws As Object
Application.DisplayAlerts = False
Set wb = Workbooks.Add(xlWBATWorksheet)
{source}{layouts}
wb.Worksheets(1).Activate
wb.SaveAs "{folder}\\pivot_show_detail.xlsx", 51
wb.Close False
Set wb = Workbooks.Open("{folder}\\pivot_show_detail.xlsx")
out = out & "@@"
For Each ws In wb.Worksheets
    If ws.PivotTables.Count > 0 Then out = out & ws.Name & "^" & Dump(ws.PivotTables(1)) & "|"
Next
wb.Close False
Build = out
End Function
'''

LAYOUT = 'out = out & Layout(wb, cache, "{name}", "{row}", "{column}", {step})\n'
REVISIONS = re.compile(r' xmlns:mc="[^"]*"| mc:Ignorable="[^"]*"| xmlns:xr="[^"]*"| xr:uid="[^"]*"| cacheId="\d+"')


def tables(path: Path) -> dict[str, str]:
    """Each pivot table's definition as the file keeps it, by the table's name."""
    with zipfile.ZipFile(path) as package:
        out: dict[str, str] = {}
        for part in sorted(name for name in package.namelist()
                           if re.fullmatch(r"xl/pivotTables/pivotTable\d+\.xml", name)):
            text = re.sub(r"^<\?xml[^>]*\?>\s*", "", package.read(part).decode("utf-8"))
            name = re.search(r'<pivotTableDefinition\b[^>]*?\bname="([^"]*)"', text)
            out[name.group(1) if name else part] = REVISIONS.sub("", text)
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
    strip_save_path(OUT / "pivot_show_detail.xlsx")
    dumps = dict(one.split("^", 1) for one in session.split("|") if "^" in one)
    again = dict(one.split("^", 1) for one in reopened.split("|") if "^" in one)
    held = tables(OUT / "pivot_show_detail.xlsx")
    record = {"build": BUILD, "source": SOURCE, "layout": LAYOUT, "steps": steps,
              "layouts": [{"name": name, "row": row, "column": column, "step": step,
                           "dump": dumps.get(name, ""), "reopened": again.get(name, ""), "file": held.get(name, "")}
                          for name, (row, column, step) in LAYOUTS.items()]}
    (OUT / "pivot_show_detail.json").write_text(json.dumps(record, indent=1, ensure_ascii=False) + "\n",
                                                encoding="utf-8")
    for layout in record["layouts"]:
        print("==", layout["name"])
        print("   session ", layout["dump"][:700])
        if layout["reopened"] != layout["dump"]:
            print("   reopened", layout["reopened"][:700])


if __name__ == "__main__":
    main()
