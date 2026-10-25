"""A pivot table's calculated items, PivotField.CalculatedItems, as Excel works them out, reads and saves them.

A macro adds an item to a field of a pivot table whose formula names the
field's other items, East+West, East*2 or a worksheet function over them,
on the rows, the columns, inside or outside another field, beside another
calculated item or one it names; changes, hides and deletes it; and tries
what Excel refuses. Every layout sits on a sheet of its own in one new
workbook, a pivot table of its own over a cache of its own, since a
calculated item belongs to its cache, over one list of regions, products,
dates, amounts and units. The dump reads every cell of the table, each
field's name and orientation and its items, a calculated one with its
formula. The workbook is saved, its pivot table and cache parts kept, and
Excel opens it and dumps every table again.

    python scripts/measure_pivot_calculated_items.py

writes tests/fixtures/pivot_calculated_items/: the workbook as Excel saved
it, and pivot_calculated_items.json, which
tests/test_excel_pivot_calculated_items.py replays.
"""

from __future__ import annotations

import json
import posixpath
import re
import shutil
import sys
import zipfile
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "scripts"))
from fixture_workbook import strip_save_path  # noqa: E402

OUT = ROOT / "tests" / "fixtures" / "pivot_calculated_items"

HELPER = '''Private note As String

Private Function Dump(pt As Object) As String
    Dim out As String, c As Object, f As Object, it As Object
    On Error Resume Next
    out = "range=" & pt.TableRange2.Address(False, False) & " cells="
    For Each c In pt.TableRange2.Cells
        If VarType(c.Value2) = vbString Then
            If IsNumeric(c.Value2) Then out = out & "'"
        End If
        out = out & CStr(c.Value2)
        If c.NumberFormat <> "General" Then out = out & "~" & c.NumberFormat
        out = out & ","
    Next
    out = out & " fields=" & pt.PivotFields.Count
    For Each f In pt.PivotFields
        out = out & " [" & f.Name & ":" & f.Orientation
        If f.Orientation = xlRowField Or f.Orientation = xlColumnField Or f.Orientation = xlPageField Then
            For Each it In f.PivotItems
                out = out & " " & it.Name
                If it.IsCalculated Then out = out & "{" & it.Formula & "}"
                If Not it.Visible Then out = out & "(hidden)"
            Next
        End If
        out = out & "]"
    Next
    Dump = out
End Function

Private Function Made(ws As Object, cache As Object, rowFields As String, columnFields As String) As Object
    Dim pt As Object, n As Variant
    Set pt = cache.CreatePivotTable(ws.Range("A5"), ws.Name)
    For Each n In Split(rowFields, ",")
        pt.PivotFields(CStr(n)).Orientation = xlRowField
    Next
    For Each n In Split(columnFields, ",")
        pt.PivotFields(CStr(n)).Orientation = xlColumnField
    Next
    pt.AddDataField pt.PivotFields("Units"), "Sum of Units", xlSum
    Set Made = pt
End Function

Private Function Layout(wb As Object, name As String, rowFields As String, columnFields As String, _
        step As Long) As String
    Dim ws As Object, pt As Object, out As String, cache As Object, ci As Object
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = name
    Set cache = wb.PivotCaches.Create(xlDatabase, "data!R1C1:R17C5")
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

SOURCE = '''Set data = wb.Worksheets(1)
data.Name = "data"
data.Range("A1:E1").Value = Array("Region", "Product", "Date", "Amount", "Units")
rows = Array("East|Apple|2024|1|15|10|1", "East|Banana|2024|2|20|20|2", "East|Cherry|2024|3|5|30|3", _
    "West|Apple|2024|5|11|40|4", "West|Banana|2024|7|30|50|5", "North|Apple|2024|8|2|60|6", _
    "North|Cherry|2024|10|19|70|7", "South|Banana|2024|12|24|80|8", "South|Cherry|2025|1|3|90|9", _
    "East|Apple|2025|2|14|15|1", "West|Apple|2025|3|31|25|2", "North|Apple|2025|4|1|35|3", _
    "South|Apple|2025|6|15|45|4", "West|Banana|2025|8|8|55|5", "South|Banana|2025|9|9|65|6", _
    "North|Cherry|2025|11|30|75|7")
For r = 0 To 15
    parts = Split(rows(r), "|")
    data.Cells(r + 2, 1).Value = parts(0)
    data.Cells(r + 2, 2).Value = parts(1)
    data.Cells(r + 2, 3).Value = DateSerial(CLng(parts(2)), CLng(parts(3)), CLng(parts(4)))
    data.Cells(r + 2, 4).Value = CLng(parts(5))
    data.Cells(r + 2, 5).Value = CLng(parts(6))
Next
'''


def probe(text: str, statement: str) -> str:
    return (f"Err.Clear\n{statement}\n"
            f'If Err.Number <> 0 Then note = note & "{text}=E" & Err.Number & ";"\nErr.Clear\n')


def note(text: str, expression: str) -> str:
    return (f'Err.Clear\nnote = note & "{text}=" & {expression} & ";"\n'
            f'If Err.Number <> 0 Then note = note & "{text}=E" & Err.Number & ";"\nErr.Clear\n')


def lines(*statements: str) -> str:
    return "\n".join(statements)


def add(field: str, name: str, formula: str, text: str = "add") -> str:
    quoted = formula.replace('"', '""')
    return probe(text, f'pt.PivotFields("{field}").CalculatedItems.Add "{name}", "{quoted}"')


def item(field: str, name: str) -> str:
    return f'pt.PivotFields("{field}").PivotItems("{name}")'


COASTS = add("Region", "Coasts", "=East+West")

#: name -> (the row fields, the column fields, what the layout does to its table).
LAYOUTS: dict[str, tuple[str, str, str]] = {
    "normal": ("Region", "", ""),
    "add_sum": ("Region", "", COASTS),
    "add_times": ("Region", "", add("Region", "Double", "=East*2")),
    "add_constant": ("Region", "", add("Region", "Plus", "=East+100")),
    "add_ratio": ("Region", "", add("Region", "Ratio", "=East/West")),
    "add_quoted": ("Region", "", add("Region", "Quoted", "='East'+'West'")),
    "add_bracket": ("Region", "", add("Region", "Bracket", "=Region[East]+Region[West]")),
    "add_function": ("Region", "", add("Region", "Avg", "=AVERAGE(East,West)")),
    "add_zero_division": ("Region", "", add("Region", "Over", "=East/0")),
    "add_columns": ("Region", "Product", COASTS),
    "add_on_columns": ("Product", "Region", COASTS),
    "add_inner": ("Product,Region", "", COASTS),
    "add_outer": ("Region,Product", "", COASTS),
    "add_two": ("Region", "", lines(COASTS, add("Region", "Double", "=East*2", "second"))),
    "add_referring": ("Region", "", lines(COASTS, add("Region", "Twice", "=Coasts*2", "second"))),
    "add_two_data": ("Region", "", lines('pt.AddDataField pt.PivotFields("Amount"), "Sum of Amount", xlSum', COASTS)),
    "add_count": ("Region", "", lines('pt.DataFields(1).Function = xlCount', COASTS)),
    "add_then_count": ("Region", "", lines(COASTS, probe("count", 'pt.DataFields(1).Function = xlCount'))),
    "add_average": ("Region", "", lines(probe("average", 'pt.DataFields(1).Function = xlAverage'), COASTS)),
    "hide": ("Region", "", lines(COASTS, probe("hide", f'{item("Region", "Coasts")}.Visible = False'))),
    "hide_member": ("Region", "", lines(COASTS, probe("hide", f'{item("Region", "East")}.Visible = False'))),
    "formula_set": ("Region", "", lines(COASTS, probe("set", f'{item("Region", "Coasts")}.Formula = "=North+South"'))),
    "delete": ("Region", "", lines(COASTS, probe("delete", f'{item("Region", "Coasts")}.Delete'))),
    "delete_member": ("Region", "", lines(COASTS, probe("delete", f'{item("Region", "East")}.Delete'))),
    "rename": ("Region", "", lines(COASTS, probe("rename", f'{item("Region", "Coasts")}.Name = "Shores"'))),
    "sorted": ("Region", "", lines(COASTS, probe("sort", 'pt.PivotFields("Region").AutoSort xlDescending, "Region"'))),
    "value_sorted": ("Region", "", lines(COASTS, probe("sort", 'pt.PivotFields("Region").AutoSort xlDescending, '
                                                               '"Sum of Units"'))),
    "page": ("Product", "", lines(COASTS, probe("page", 'pt.PivotFields("Region").Orientation = xlPageField'),
                                  probe("choose", 'pt.PivotFields("Region").CurrentPage = "Coasts"'))),
    "hidden_field": ("Product", "", lines(COASTS, probe("off", 'pt.PivotFields("Region").Orientation = xlHidden'))),
    "refresh": ("Region", "", lines(COASTS, probe("refresh", "pt.PivotCache.Refresh"))),
    "collection": ("Region", "", lines(
        COASTS, add("Region", "Double", "=East*2", "second"),
        note("count", 'pt.PivotFields("Region").CalculatedItems.Count'),
        note("first", 'pt.PivotFields("Region").CalculatedItems(1).Name'),
        note("by_name", 'pt.PivotFields("Region").CalculatedItems("Double").Formula'),
        note("source", f'{item("Region", "Coasts")}.SourceName'),
        note("value", f'{item("Region", "Coasts")}.Value'),
        note("member", f'{item("Region", "East")}.IsCalculated'),
        note("member_formula", f'{item("Region", "East")}.Formula'),
        note("position", f'{item("Region", "Coasts")}.Position'),
        note("other_count", 'pt.PivotFields("Product").CalculatedItems.Count'),
        note("missing", 'pt.PivotFields("Region").CalculatedItems(3).Name'),
        note("standard", f'{item("Region", "Coasts")}.StandardFormula'))),
    "standard": ("Region", "", probe("add", 'pt.PivotFields("Region").CalculatedItems.Add "Coasts", '
                                            '"=East+West", True')),
    "errors": ("Region", "", lines(
        add("Region", "East", "=West*2", "existing"),
        add("Region", "Nowhere", "=Nowhere+1", "unknown"),
        add("Region", "", "=East+West", "empty"),
        add("Region", "Text", "=East&West", "text"),
        add("Region", "Field", "=Units*2", "field"),
        add("Region", "Product", "=Apple+1", "other_field"),
        probe("data", 'pt.DataFields(1).CalculatedItems.Add "Data", "=East+West"'),
        probe("values", 'pt.DataPivotField.CalculatedItems.Add "Data", "=East+West"'))),
    "errors_page": ("Product", "", lines(probe("page", 'pt.PivotFields("Region").Orientation = xlPageField'),
                                         COASTS)),
    "errors_hidden_field": ("Product", "", COASTS),
    "errors_calculated_field": ("Region", "", lines('pt.CalculatedFields.Add "Twice", "=Units*2"',
                                                     'pt.AddDataField pt.PivotFields("Twice"), "Sum of Twice", xlSum',
                                                     COASTS)),
    # Round two: totals by other functions and of a formula that is not a sum, a page field, names that need
    # quoting, numbers, and another table of the cache.
    **{f"function_{name}": ("Region", "Product", lines(probe("function", f"pt.DataFields(1).Function = {constant}"),
                                                       COASTS))
       for name, constant in (("max", "xlMax"), ("min", "xlMin"), ("product", "xlProduct"),
                              ("count_numbers", "xlCountNums"), ("stdev", "xlStDev"), ("var", "xlVar"))},
    "function_after": ("Region", "", lines(COASTS, probe("average", "pt.DataFields(1).Function = xlAverage"),
                                           probe("max", "pt.DataFields(1).Function = xlMax"))),
    "ratio_columns": ("Region", "Product", add("Region", "Ratio", "=East/West")),
    "ratio_outer": ("Region,Product", "", add("Region", "Ratio", "=East/West")),
    "ratio_inner": ("Product,Region", "", add("Region", "Ratio", "=East/West")),
    "page_after": ("Product", "", lines(add("Region", "Coasts", "=East+West"),
                                        probe("move", 'pt.PivotFields("Region").Orientation = xlPageField'),
                                        probe("choose", 'pt.PivotFields("Region").CurrentPage = "Coasts"'))),
    "page_rows": ("Region", "", lines(COASTS, probe("move", 'pt.PivotFields("Region").Orientation = xlPageField'),
                                      probe("rows", 'pt.PivotFields("Product").Orientation = xlRowField'))),
    "columns_after": ("Region", "", lines(COASTS, probe("move", 'pt.PivotFields("Region").Orientation = '
                                                                 'xlColumnField'))),
    "returned": ("Region", "", lines('Set ci = pt.PivotFields("Region").CalculatedItems.Add("Coasts", "=East+West")',
                                     note("name", "ci.Name"), note("parent", "ci.Parent.Name"),
                                     note("calculated", "ci.IsCalculated"))),
    "other_field_item": ("Region", "", add("Region", "Mixed", "=Product[Apple]+1")),
    "position_reference": ("Region", "", add("Region", "First", "=Region[1]*2")),
    "numbers": ("Amount", "", add("Amount", "Pair", "='10'+'20'")),
    "numbers_plain": ("Amount", "", add("Amount", "Pair", "=10+20")),
    "odd_names": ("Region", "", lines(
        'wb.Worksheets("data").Range("A2").Value = "North East"', 'pt.PivotCache.Refresh',
        add("Region", "Spaced", "='North East'*2"),
        'wb.Worksheets("data").Range("A2").Value = "East"')),
    "shared_cache": ("Region", "", lines(
        'Set ci = cache.CreatePivotTable(ws.Range("H5"), name & "_2")',
        'ci.PivotFields("Region").Orientation = xlRowField',
        'ci.AddDataField ci.PivotFields("Units"), "Sum of Units", xlSum',
        COASTS, 'note = note & "second=" & Dump(ci) & ";"')),
    "shared_cache_hidden": ("Region", "", lines(
        'Set ci = cache.CreatePivotTable(ws.Range("H5"), name & "_2")',
        'ci.PivotFields("Product").Orientation = xlRowField',
        'ci.AddDataField ci.PivotFields("Units"), "Sum of Units", xlSum',
        COASTS, 'ci.PivotFields("Region").Orientation = xlColumnField',
        'note = note & "second=" & Dump(ci) & ";"')),
    "hide_referenced_both": ("Region", "", lines(COASTS, probe("east", f'{item("Region", "East")}.Visible = False'),
                                                 probe("west", f'{item("Region", "West")}.Visible = False'))),
    "filtered_other": ("Region", "Product", lines(COASTS, probe("apple", f'{item("Product", "Apple")}.Visible = '
                                                                          'False'))),
    "page_filtered": ("Region", "", lines(COASTS, probe("page", 'pt.PivotFields("Product").Orientation = '
                                                                 'xlPageField'),
                                          probe("choose", 'pt.PivotFields("Product").CurrentPage = "Apple"'))),
    # Round three: where a calculated item's lines go where the items it names have none, constants, references
    # between items, how a formula is spelled back, and grouping a field that has one.
    "absent_inner": ("Product,Region", "", add("Region", "W2", "=West*2")),
    "absent_columns": ("Product", "Region", add("Region", "W2", "=West*2")),
    "absent_outer": ("Region,Product", "", add("Region", "W2", "=West*2")),
    "constant_inner": ("Product,Region", "", add("Region", "Ten", "=10")),
    "constant_columns": ("Region", "Product", add("Region", "Ten", "=10")),
    "hidden_referenced": ("Region", "", lines(COASTS, add("Region", "Twice", "=Coasts*2", "second"),
                                              probe("hide", f'{item("Region", "Coasts")}.Visible = False'))),
    "delete_referenced": ("Region", "", lines(COASTS, add("Region", "Twice", "=Coasts*2", "second"),
                                              probe("delete", f'{item("Region", "Coasts")}.Delete'))),
    "spelled": ("Region", "", lines(add("Region", "Spaces", "= East + West ", "spaces"),
                                    add("Region", "Lower", "=sum(east,west)", "lower"),
                                    add("Region", "Number", "=East*0.50+1E1", "number"))),
    "group_numbers_with_item": ("Amount", "", lines(add("Amount", "Pair", "='10'+'20'"),
                                                    probe("group", 'pt.PivotFields("Amount").DataRange.Cells(1).Group '
                                                                   'Start:=0, End:=100, By:=10'))),
    "subtotal_function": ("Product,Region", "", lines(
        COASTS, probe("subtotals", 'pt.PivotFields("Product").Subtotals = Array(False, False, False, False, '
                                   'True, False, False, False, False, False, False, False)'))),
}

BUILD = HELPER + '''
Public Function Build() As String
Dim wb As Object, data As Object, out As String, r As Long, rows As Variant, parts As Variant
Dim ws As Object
Application.DisplayAlerts = False
Set wb = Workbooks.Add(xlWBATWorksheet)
{source}{layouts}
wb.Worksheets(1).Activate
wb.SaveAs "{folder}\\pivot_calculated_items.xlsx", 51
wb.Close False
Set wb = Workbooks.Open("{folder}\\pivot_calculated_items.xlsx")
out = out & "@@"
For Each ws In wb.Worksheets
    If ws.PivotTables.Count > 0 Then out = out & ws.Name & "^" & Dump(ws.PivotTables(ws.Name)) & "|"
Next
wb.Close False
Build = out
End Function
'''

LAYOUT = 'out = out & Layout(wb, "{name}", "{row}", "{column}", {step})\n'
REVISIONS = re.compile(r' xmlns:mc="[^"]*"| mc:Ignorable="[^"]*"| xmlns:xr="[^"]*"| xr:uid="[^"]*"| cacheId="\d+"'
                       r'| refreshedDate="[^"]*"| refreshedBy="[^"]*"| r:id="[^"]*"')


def tables(path: Path) -> dict[str, str]:
    """Each pivot table's definition and its cache's definition, as the file keeps them, by the table's name."""
    with zipfile.ZipFile(path) as package:
        out: dict[str, str] = {}
        for part in sorted(name for name in package.namelist()
                           if re.fullmatch(r"xl/pivotTables/pivotTable\d+\.xml", name)):
            text = package.read(part).decode("utf-8")
            rels = package.read(posixpath.join("xl/pivotTables/_rels", posixpath.basename(part) + ".rels"))
            target = re.search(rb'Target="([^"]*pivotCacheDefinition\d+\.xml)"', rels)
            cache = ""
            if target:
                cache = package.read(posixpath.normpath(posixpath.join("xl/pivotTables",
                                                                       target.group(1).decode()))).decode("utf-8")
            name = re.search(r'<pivotTableDefinition\b[^>]*?\bname="([^"]*)"', text)
            body = re.sub(r"^<\?xml[^>]*\?>\s*", "", text)
            fields = re.search(r"<cacheFields\b.*?</cacheFields>(?:<calculatedItems\b.*?</calculatedItems>)?", cache,
                               re.DOTALL)
            out[name.group(1) if name else part] = REVISIONS.sub("", body + (fields.group() if fields else ""))
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
    strip_save_path(OUT / "pivot_calculated_items.xlsx")
    dumps = dict(one.split("^", 1) for one in session.split("|") if "^" in one)
    again = dict(one.split("^", 1) for one in reopened.split("|") if "^" in one)
    held = tables(OUT / "pivot_calculated_items.xlsx")
    record = {"build": BUILD, "source": SOURCE, "layout": LAYOUT, "steps": steps,
              "layouts": [{"name": name, "row": row, "column": column, "step": step,
                           "dump": dumps.get(name, ""), "reopened": again.get(name, ""), "file": held.get(name, "")}
                          for name, (row, column, step) in LAYOUTS.items()]}
    (OUT / "pivot_calculated_items.json").write_text(json.dumps(record, indent=1, ensure_ascii=False) + "\n",
                                                     encoding="utf-8")
    for layout in record["layouts"]:
        print("==", layout["name"])
        print("   session ", layout["dump"][:900])
        if layout["reopened"] != layout["dump"]:
            print("   reopened", layout["reopened"][:900])


if __name__ == "__main__":
    main()
