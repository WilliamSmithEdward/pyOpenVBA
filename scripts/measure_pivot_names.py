"""A pivot table's data fields, fields and items renamed, as Excel takes, shows and saves the names.

A macro sets the Caption, Name or Value of a data field, of the Values
field, of a field on an axis or on none, and of an item. Every layout sits
on a sheet of its own in one new workbook, a pivot table of its own over
one source list of regions, products, channels, amounts and units:
regions on the rows, products on the columns, the sum of the amounts.
The dump reads every cell of the table, each data field's name, caption,
source name and function, each field's name, caption and source name,
and each region's name, caption, source name and value. The workbook is
saved, its pivot table parts kept, and Excel opens it and dumps every
table again.

    python scripts/measure_pivot_names.py

writes tests/fixtures/pivot_names/: the workbook as Excel saved it, and
pivot_names.json, which tests/test_excel_pivot_names.py replays.
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

OUT = ROOT / "tests" / "fixtures" / "pivot_names"

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
    Dim out As String, c As Object, f As Object, it As Object
    On Error Resume Next
    out = "range=" & pt.TableRange2.Address(False, False) & " cells="
    For Each c In pt.TableRange2.Cells
        out = out & CStr(c.Value2) & ","
    Next
    out = out & " values=" & pt.DataPivotField.Name & "/" & pt.DataPivotField.Caption
    For Each f In pt.DataFields
        out = out & " " & f.Name & ":" & Props(f, "Caption SourceName Function")
    Next
    For Each f In pt.PivotFields
        out = out & " [" & f.Name & ":" & Props(f, "Caption SourceName Orientation") & "]"
    Next
    For Each it In pt.PivotFields(1).PivotItems
        out = out & " {" & it.Name & ":" & Props(it, "Caption SourceName Value Position") & "}"
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
    Dim ws As Object, pt As Object, out As String, total As Object
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


def note(text: str, expression: str) -> str:
    return (f'Err.Clear\nnote = note & "{text}=" & {expression} & ";"\n'
            f'If Err.Number <> 0 Then note = note & "{text}=E" & Err.Number & ";"\nErr.Clear\n')


UNITS = 'pt.AddDataField pt.PivotFields("Units"), "Sum of Units", xlSum'
EAST = 'pt.PivotFields("Region").PivotItems("East")'
OUTLINE = "pt.RowAxisLayout xlOutlineRow"

#: name -> (the row fields, the column fields, what the layout does to its table).
LAYOUTS: dict[str, tuple[str, str, str]] = {
    "normal": ("Region", "Product", ""),
    # A data field's name.
    "data_caption": ("Region", "Product", probe("caption", 'total.Caption = "Total Sales"')),
    "data_name": ("Region", "Product", probe("name", 'total.Name = "Total Sales"')),
    "data_source_name": ("Region", "Product", lines(probe("same", 'total.Caption = "Amount"'),
                                                    probe("case", 'total.Caption = "amount"'),
                                                    probe("other_field", 'total.Caption = "Region"'),
                                                    probe("spaced", 'total.Caption = "Amount "'))),
    "data_duplicate": ("Region", "Product", lines(UNITS, probe("duplicate", 'pt.DataFields(2).Caption = "Sum of '
                                                                            'Amount"'),
                                                  probe("duplicate_case", 'pt.DataFields(2).Caption = "sum of '
                                                                          'amount"'),
                                                  probe("item", 'pt.DataFields(2).Caption = "East"'))),
    "data_empty": ("Region", "Product", probe("empty", 'total.Caption = ""')),
    "data_long": ("Region", "Product", lines(probe("long", 'total.Caption = String(256, "x")'),
                                             probe("most", 'total.Caption = String(255, "y")'))),
    "data_two": ("Region", "Product", lines(UNITS, probe("second", 'pt.DataFields(2).Caption = "Units Sold"'))),
    "data_two_rows": ("Region", "Product", lines(UNITS, "pt.DataPivotField.Orientation = xlRowField",
                                                 probe("second", 'pt.DataFields(2).Caption = "Units Sold"'))),
    "data_then_function": ("Region", "Product", lines(probe("caption", 'total.Caption = "Sales"'),
                                                      probe("function", "total.Function = xlCount"))),
    "data_lookup": ("Region", "Product", lines(probe("caption", 'total.Caption = "Sales"'),
                                               note("by_data", 'pt.DataFields("Sales").Function'),
                                               note("by_field", 'pt.PivotFields("Sales").Function'),
                                               note("old", 'pt.DataFields("Sum of Amount").Function'))),
    "data_outline": ("Region", "Product", lines(OUTLINE, probe("caption", 'total.Caption = "Sales"'))),
    "data_no_axes": ("", "", probe("caption", 'total.Caption = "Sales"')),
    # The Values field's name.
    "values_caption": ("Region", "Product", lines(UNITS, probe("caption", 'pt.DataPivotField.Caption = "Measures"'))),
    "values_name": ("Region", "Product", lines(UNITS, probe("name", 'pt.DataPivotField.Name = "Measures"'))),
    "values_outline": ("Region", "Product", lines(UNITS, OUTLINE, "pt.DataPivotField.Orientation = xlRowField",
                                                  probe("caption", 'pt.DataPivotField.Caption = "Measures"'))),
    # A field's name, on an axis or on none.
    "row_outline": ("Region", "Product", lines(OUTLINE, probe("caption", 'pt.PivotFields("Region").Caption = '
                                                                         '"Area"'))),
    "row_name": ("Region", "Product", lines(OUTLINE, probe("name", 'pt.PivotFields("Region").Name = "Area"'))),
    "hidden_name": ("Region", "Product", lines(probe("name", 'pt.PivotFields("Channel").Name = "Route"'),
                                               note("lookup", 'pt.PivotFields("Route").SourceName'))),
    "hidden_caption": ("Region", "Product", probe("caption", 'pt.PivotFields("Channel").Caption = "Route"')),
    "field_duplicate": ("Region", "Product", lines(probe("product", 'pt.PivotFields("Region").Caption = "Product"'),
                                                   probe("data", 'pt.PivotFields("Region").Caption = "Sum of '
                                                                 'Amount"'),
                                                   probe("case", 'pt.PivotFields("Region").Caption = "region"'),
                                                   probe("empty", 'pt.PivotFields("Region").Caption = ""'))),
    "field_page": ("Region", "Product", lines('pt.PivotFields("Channel").Orientation = xlPageField',
                                              probe("caption", 'pt.PivotFields("Channel").Caption = "Route"'))),
    # An item's name.
    "item_caption": ("Region", "Product", probe("caption", f'{EAST}.Caption = "Eastern"')),
    "item_name": ("Region", "Product", probe("name", f'{EAST}.Name = "Eastern"')),
    "item_value": ("Region", "Product", probe("value", f'{EAST}.Value = "Eastern"')),
    "item_duplicate": ("Region", "Product", lines(probe("duplicate", f'{EAST}.Caption = "West"'),
                                                  probe("duplicate_case", f'{EAST}.Caption = "west"'),
                                                  probe("case", f'{EAST}.Caption = "EAST"'))),
    "item_lookup": ("Region", "Product", lines(probe("caption", f'{EAST}.Caption = "Eastern"'),
                                               note("new", 'pt.PivotFields("Region").PivotItems("Eastern").SourceName'),
                                               note("old", 'pt.PivotFields("Region").PivotItems("East").Name'))),
    "item_reset": ("Region", "Product", lines(probe("caption", f'{EAST}.Caption = "Eastern"'),
                                              probe("reset", 'pt.PivotFields("Region").PivotItems("Eastern").Caption '
                                                             '= ""'))),
    "item_back": ("Region", "Product", lines(probe("caption", f'{EAST}.Caption = "Eastern"'),
                                             probe("back", 'pt.PivotFields("Region").PivotItems("Eastern").Caption = '
                                                           '"East"'))),
    "item_sorted": ("Region", "Product", lines(probe("caption", 'pt.PivotFields("Region").PivotItems("West").Caption '
                                                                '= "Atlantic"'),
                                               'pt.PivotFields("Region").AutoSort xlAscending, "Region"')),
    "item_then_sorted": ("Region", "Product", lines('pt.PivotFields("Region").AutoSort xlAscending, "Region"',
                                                    probe("caption", 'pt.PivotFields("Region").PivotItems("West")'
                                                                     '.Caption = "Atlantic"'))),
    "item_page": ("Channel", "Product", lines('pt.PivotFields("Region").Orientation = xlPageField',
                                              probe("caption", f'{EAST}.Caption = "Eastern"'),
                                              probe("page", 'pt.PivotFields("Region").CurrentPage = "Eastern"'))),
    "item_numbers": ("Units", "Product", lines(probe("caption", 'pt.PivotFields("Units").PivotItems("1").Caption = '
                                                                '"One"'),
                                               note("lookup", 'pt.PivotFields("Units").PivotItems("One").Value'))),
    "item_column": ("Region", "Product", probe("caption", 'pt.PivotFields("Product").PivotItems("Apple").Caption = '
                                                          '"Apples"')),
    "item_label_filter": ("Region", "Product", lines(probe("caption", f'{EAST}.Caption = "Eastern"'),
                                                     probe("filter", 'pt.PivotFields("Region").PivotFilters.Add2 '
                                                                     'xlCaptionBeginsWith, , "Ea"'))),
    "item_hidden": ("Region", "Product", lines(probe("caption", f'{EAST}.Caption = "Eastern"'),
                                               probe("hide", 'pt.PivotFields("Region").PivotItems("Eastern").Visible '
                                                             '= False'))),
    # A long name, names a caption takes, an empty Name, and the items of a field on no axis.
    "data_longest": ("Region", "Product", lines(probe("long", 'total.Caption = String(256, "x")'),
                                                note("length", "Len(total.Caption)"))),
    "data_field_caption": ("Region", "Product", lines('pt.PivotFields("Region").Caption = "Area"',
                                                      probe("area", 'total.Caption = "Area"'),
                                                      probe("region", 'total.Caption = "Region"'))),
    "field_empty_name": ("Region", "Product", lines(OUTLINE, probe("name", 'pt.PivotFields("Region").Name = ""'))),
    "values_field_name": ("Region", "Product", lines(UNITS, probe("region", 'pt.DataPivotField.Caption = "Region"'),
                                                     probe("data", 'pt.DataPivotField.Caption = "Sum of Units"'))),
    "item_off_axis": ("Product", "", lines(probe("caption", f'{EAST}.Caption = "Eastern"'),
                                           note("listed", 'pt.PivotFields("Region").PivotItems(1).Name'))),
    "item_get_data": ("Region", "Product", lines(probe("caption", f'{EAST}.Caption = "Eastern"'),
                                                 note("new", 'pt.GetPivotData("Sum of Amount", "Region", "Eastern")'),
                                                 note("old", 'pt.GetPivotData("Sum of Amount", "Region", "East")'))),
}

BUILD = HELPER + '''
Public Function Build() As String
Dim wb As Object, data As Object, cache As Object, out As String, r As Long, rows As Variant, parts As Variant
Dim ws As Object
Application.DisplayAlerts = False
Set wb = Workbooks.Add(xlWBATWorksheet)
{source}{layouts}
wb.Worksheets(1).Activate
wb.SaveAs "{folder}\\pivot_names.xlsx", 51
wb.Close False
Set wb = Workbooks.Open("{folder}\\pivot_names.xlsx")
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
    strip_save_path(OUT / "pivot_names.xlsx")
    dumps = dict(one.split("^", 1) for one in session.split("|") if "^" in one)
    again = dict(one.split("^", 1) for one in reopened.split("|") if "^" in one)
    held = tables(OUT / "pivot_names.xlsx")
    record = {"build": BUILD, "source": SOURCE, "layout": LAYOUT, "steps": steps,
              "layouts": [{"name": name, "row": row, "column": column, "step": step,
                           "dump": dumps.get(name, ""), "reopened": again.get(name, ""), "file": held.get(name, "")}
                          for name, (row, column, step) in LAYOUTS.items()]}
    (OUT / "pivot_names.json").write_text(json.dumps(record, indent=1, ensure_ascii=False) + "\n", encoding="utf-8")
    for layout in record["layouts"]:
        print("==", layout["name"])
        print("   session ", layout["dump"][:600])
        if layout["reopened"] != layout["dump"]:
            print("   reopened", layout["reopened"][:600])


if __name__ == "__main__":
    main()
