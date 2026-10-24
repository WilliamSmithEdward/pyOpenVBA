"""A pivot table's calculated fields, as Excel works them out, reads and saves them.

A macro adds calculated fields with PivotTable.CalculatedFields.Add, puts
them in the data area, reads and changes their Formula and deletes them.
Every layout sits on a sheet of its own in one new workbook, a pivot
table of its own over a cache of its own, since a calculated field
belongs to its cache: one source list of regions, products, channels,
amounts, units and unit costs, regions on the rows, products on the
columns, the sum of the amounts. The dump reads every cell of the table
with its number format, each data field's name, function, number format,
source and calculation, and each calculated field's name, formula and
place. The workbook is saved, its pivot table and cache parts kept, and
Excel opens it and dumps every table again.

    python scripts/measure_pivot_calculated_fields.py

writes tests/fixtures/pivot_calculated_fields/: the workbook as Excel
saved it, and pivot_calculated_fields.json, which
tests/test_excel_pivot_calculated_fields.py replays.
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

OUT = ROOT / "tests" / "fixtures" / "pivot_calculated_fields"

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
        out = out & " " & f.Name & ":" & Props(f, "Function NumberFormat SourceName Calculation Formula IsCalculated")
    Next
    out = out & " calculated=" & pt.CalculatedFields.Count
    For Each f In pt.CalculatedFields
        out = out & " " & f.Name & ":" & Props(f, "Formula IsCalculated Orientation SourceName Caption")
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

Private Function Layout(wb As Object, name As String, rowFields As String, columnFields As String, _
        step As Long) As String
    Dim ws As Object, pt As Object, out As String, total As Object, cache As Object
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = name
    Set cache = wb.PivotCaches.Create(xlDatabase, "data!R1C1:R17C6")
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

#: The source list of tests/fixtures/pivot_calculations, with a unit cost beside each row.
SOURCE = '''Set data = wb.Worksheets(1)
data.Name = "data"
data.Range("A1:F1").Value = Array("Region", "Product", "Channel", "Amount", "Units", "Unit Cost")
rows = Array("East|Apple|Online|10|1|4", "East|Banana|Store|20|2|3", "East|Cherry|Online|30|3|5", _
    "West|Apple|Store|40|4|4", "West|Banana|Online|50|5|2", "North|Apple|Online|60|6|6", "North|Cherry|Store|70|7|5", _
    "South|Banana|Store|80|8|3", "South|Cherry|Online|90|9|4", "East|Apple|Store|15|1|7", "West|Apple|Online|25|2|3", _
    "North|Apple|Store|35|3|2", "South|Apple|Store|45|4|6", "West|Banana|Store|55|5|1", "South|Banana|Online|65|6|8", _
    "North|Cherry|Online|75|7|2")
For r = 0 To 15
    parts = Split(rows(r), "|")
    data.Cells(r + 2, 1).Value = parts(0)
    data.Cells(r + 2, 2).Value = parts(1)
    data.Cells(r + 2, 3).Value = parts(2)
    data.Cells(r + 2, 4).Value = CLng(parts(3))
    data.Cells(r + 2, 5).Value = CLng(parts(4))
    data.Cells(r + 2, 6).Value = CLng(parts(5))
Next
'''


def probe(text: str, statement: str) -> str:
    return (f"Err.Clear\n{statement}\n"
            f'If Err.Number <> 0 Then note = note & "{text}=E" & Err.Number & ";"\nErr.Clear\n')


def lines(*statements: str) -> str:
    return "\n".join(statements)


def calc(name: str, formula: str, *, data: bool = True) -> str:
    """A calculated field added, and put in the data area where ``data``, each statement's error noted."""
    formula = formula.replace('"', '""')
    out = [probe("add", f'pt.CalculatedFields.Add "{name}", "{formula}"')]
    if data:
        out.append(probe("data", f'pt.AddDataField pt.PivotFields("{name}")'))
    return lines(*out)


def note(text: str, expression: str) -> str:
    """A value read into the note, its error noted instead."""
    return (f'Err.Clear\nnote = note & "{text}=" & {expression} & ";"\n'
            f'If Err.Number <> 0 Then note = note & "{text}=E" & Err.Number & ";"\nErr.Clear\n')


TWO = "Region,Product"
MARGIN = calc("Margin", "=Amount*0.1")
RATIO = calc("Ratio", "=Amount/Units")

#: name -> (the row fields, the column fields, what the layout does to its table).
LAYOUTS: dict[str, tuple[str, str, str]] = {
    "normal": ("Region", "Product", ""),
    "margin": ("Region", "Product", MARGIN),
    "ratio": ("Region", "Product", RATIO),
    "spread": ("Region", "Product", calc("Spread", "=Amount-Units")),
    "constant": ("Region", "Product", calc("Five", "=5")),
    "condition": ("Region", "Product", calc("Big", "=IF(Amount>100,1,0)")),
    "rounded": ("Region", "Product", calc("Price", "=ROUND(Amount/Units,1)")),
    "quoted": ("Region", "Product", calc("Twice", "='Amount'*2")),
    "spaced": ("Region", "Product", calc("Cost", "='Unit Cost'*Units")),
    "spaced_loose": ("Region", "Product", calc("Cost", "=  'Unit Cost' *  Units")),
    "lowercase": ("Region", "Product", calc("Margin", "=amount*0.1")),
    "function_case": ("Region", "Product", calc("Root", "=sqrt(Amount)")),
    "nested": ("Region", "Product", lines(MARGIN, calc("Double", "=Margin*2"))),
    "divide_zero": ("Region", "Product", calc("Broken", "=Amount/0")),
    "text": ("Region", "Product", calc("Size", '=IF(Amount>100,"big","small")')),
    "only_calculated": ("Region", "Product", lines(MARGIN, 'pt.DataFields("Sum of Amount").Orientation = xlHidden')),
    "formula_change": ("Region", "Product", lines(MARGIN, probe("change",
                                                                'pt.PivotFields("Margin").Formula = "=Amount*0.2"'))),
    "deleted": ("Region", "Product", lines(MARGIN, probe("delete", 'pt.PivotFields("Margin").Delete'))),
    "deleted_unused": ("Region", "Product", lines(calc("Margin", "=Amount*0.1", data=False),
                                                  probe("delete", 'pt.PivotFields("Margin").Delete'))),
    "unused": ("Region", "Product", calc("Margin", "=Amount*0.1", data=False)),
    "errors": ("Region", "Product", lines(
        probe("duplicate", 'pt.CalculatedFields.Add "Amount", "=Units*2"'),
        probe("duplicate_case", 'pt.CalculatedFields.Add "amount", "=Units*2"'),
        probe("bad", 'pt.CalculatedFields.Add "Bad", "=Amount*"'),
        probe("unknown", 'pt.CalculatedFields.Add "Unknown", "=Nope*2"'),
        probe("cell", 'pt.CalculatedFields.Add "Cell", "=Amount*A1"'),
        MARGIN,
        probe("row", 'pt.PivotFields("Margin").Orientation = xlRowField'),
        probe("page", 'pt.PivotFields("Margin").Orientation = xlPageField'),
        probe("count", 'pt.DataFields("Sum of Margin").Function = xlCount'),
        probe("again", 'pt.CalculatedFields.Add "Margin", "=Amount"'),
        probe("bad_change", 'pt.PivotFields("Margin").Formula = "=Amount*"'),
        probe("source_formula", 'pt.PivotFields("Amount").Formula = "=Units"'),
        note("source_read", 'pt.PivotFields("Amount").Formula'))),
    "average": ("Region", "Product", lines(MARGIN, probe("average",
                                                         'pt.DataFields("Sum of Margin").Function = xlAverage'))),
    "two_calculated": ("Region", "Product", lines(MARGIN, RATIO)),
    "hidden_item": ("Region", "Product", lines('pt.PivotFields("Product").PivotItems("Banana").Visible = False',
                                               RATIO)),
    "page_filtered": ("Region", "Product", lines('pt.PivotFields("Channel").Orientation = xlPageField',
                                                 'pt.PivotFields("Channel").CurrentPage = "Online"', RATIO)),
    "two_rows": (TWO, "", RATIO),
    "two_rows_channel": (TWO, "Channel", RATIO),
    "count_subtotals": (TWO, "", lines('pt.PivotFields("Region").Subtotals(3) = True', RATIO)),
    "percent_of_total": ("Region", "Product", lines(MARGIN, 'Set total = pt.DataFields("Sum of Margin")',
                                                    probe("calculation", "total.Calculation = xlPercentOfTotal"))),
    "caption": ("Region", "Product", lines(probe("add", 'pt.CalculatedFields.Add "Margin", "=Amount*0.1"'),
                                           probe("data", 'pt.AddDataField pt.PivotFields("Margin"), "Margin Total", '
                                                         'xlSum'))),
    "values_rows": ("Region", "Product", lines(RATIO, "pt.DataPivotField.Orientation = xlRowField")),
    "listed": ("Region", "Product", lines(calc("Margin", "=Amount*0.1", data=False),
                                          calc("Ratio", "=Amount/Units", data=False),
                                          note("first", "pt.CalculatedFields(1).Name"),
                                          note("by_name", 'pt.CalculatedFields("Ratio").Formula'),
                                          note("count", "pt.CalculatedFields.Count"),
                                          note("fields", "pt.PivotFields.Count"),
                                          note("hidden", "pt.HiddenFields.Count"))),
    "number_format": ("Region", "Product", lines(MARGIN, probe("format", 'pt.DataFields("Sum of Margin").NumberFormat '
                                                                         '= "0.0"'))),
    "standard_formula": ("Region", "Product", lines(
        probe("add", 'pt.CalculatedFields.Add "Margin", "=Amount*0.1", True'),
        probe("data", 'pt.AddDataField pt.PivotFields("Margin")'),
        note("standard", 'pt.PivotFields("Margin").StandardFormula'))),
    "by_orientation": ("Region", "Product", lines(probe("add", 'pt.CalculatedFields.Add "Margin", "=Amount*0.1"'),
                                                  probe("orient", 'pt.PivotFields("Margin").Orientation = '
                                                                  'xlDataField'))),
    "returned_field": ("Region", "Product", lines(
        probe("add", 'Set total = pt.CalculatedFields.Add("Margin", "=Amount*0.1")'),
        note("returned", 'total.Name & "/" & total.Formula & "/" & total.Orientation'))),
    # How Excel spells a formula back: spaces, literals, operators, references and functions.
    "spellings": ("Region", "Product", lines(*(calc(f"S{index:02}", formula, data=False) for index, formula in enumerate((
        "=Amount * Units", "=Amount  *Units", "=Amount*  Units", "= Amount*Units", "=Amount *Units",
        "=ROUND( Amount ,1)", "=ROUND(Amount, 1)", "=(Amount + Units)*2", "=Amount* -Units", "=- Amount",
        "=Amount*0.10", "=Amount*1E3", "=Amount/100%", "=Amount^2", "=Amount<>Units", "=IF(TRUE,Amount,0)",
        "=Amount&Units", "=((Amount))", "=round(amount,1)", "=Amount*Units ", "=Amount*.5", "=Amount*5.0E-1",
        "=Amount>=Units", "=IF(Amount>Units,Amount,Units)", "=MAX(Amount,Units)", "=AMOUNT", "='amount'",
        "='Unit Cost'", "=Unit_Cost", "=Amount/'Unit Cost'", "=  Amount", "=Amount+0", "=+Amount",
        "=Amount*1/3", "=SUM(Amount,Units)", "=Amount--Units", "=-(-Amount)", "=Amount*2.50"), 1)))),
    "operators": ("Region", "Product", lines(
        calc("Power", "=Amount^2"), calc("Unequal", "=Amount<>Units"), calc("Joined", "=Amount&Units"),
        calc("Negated", "=-Amount"), calc("Percent", "=Amount/100%"), calc("Larger", "=MAX(Amount,Units)"),
        'pt.DataFields("Sum of Amount").Orientation = xlHidden')),
    "text_field": ("Region", "Product", calc("Regions", "=Region")),
    "text_field_count": ("Region", "Product", calc("Counted", "=Region*2")),
    "spaced_name": ("Region", "Product", lines(calc("Unit Margin", "=Amount*0.1"), calc("Twice", "='Unit Margin'*2"))),
    "renamed": ("Region", "Product", lines(MARGIN, calc("Double", "=Margin*2", data=False),
                                           probe("rename", 'pt.PivotFields("Margin").Name = "Profit"'))),
    "renamed_caption": ("Region", "Product", lines(MARGIN, probe("rename", 'pt.PivotFields("Margin").Caption = '
                                                                           '"Profit"'))),
    "delete_referenced": ("Region", "Product", lines(MARGIN, calc("Double", "=Margin*2"),
                                                     probe("delete", 'pt.PivotFields("Margin").Delete'))),
    "delete_by_collection": ("Region", "Product", lines(MARGIN, probe("delete", 'pt.CalculatedFields("Margin").Delete'))),
    "circular": ("Region", "Product", lines(calc("Loop", "=Loop*2"), calc("Margin", "=Amount*0.1"),
                                            probe("circular", 'pt.PivotFields("Margin").Formula = "=Margin+1"'))),
    "sort_by_calculated": ("Region", "Product", lines(RATIO, probe("sort", 'pt.PivotFields("Region").AutoSort '
                                                                          'xlDescending, "Sum of Ratio"'))),
    "items_of_calculated": ("Region", "Product", lines(MARGIN, note("items", 'pt.PivotFields("Margin").PivotItems.Count'),
                                                       note("data_type", 'pt.PivotFields("Margin").DataType'),
                                                       note("source", 'pt.PivotFields("Margin").SourceName'))),
    "value_filter": ("Region", "Product", lines(RATIO, probe("filter", 'pt.PivotFields("Region").PivotFilters.Add2 '
                                                                       'xlValueIsGreaterThan, pt.DataFields("Sum of '
                                                                       'Ratio"), 10.5'))),
    "difference_from": ("Region", "Product", lines(MARGIN, 'Set total = pt.DataFields("Sum of Margin")',
                                                   probe("calculation", "total.Calculation = xlDifferenceFrom"),
                                                   probe("field", 'total.BaseField = "Region"'),
                                                   probe("item", 'total.BaseItem = "East"'))),
    "running_total": ("Region", "Product", lines(RATIO, 'Set total = pt.DataFields("Sum of Ratio")',
                                                 probe("calculation", "total.Calculation = xlRunningTotal"),
                                                 probe("field", 'total.BaseField = "Region"'))),
    "page_area": (TWO, "", lines('pt.PivotFields("Channel").Orientation = xlPageField', MARGIN)),
    # Where the spaces in a formula go.
    "spacing": ("Region", "Product", lines(*(calc(f"T{index:02}", formula, data=False) for index, formula in enumerate((
        "= Amount*Units*2", "= Amount+Units*2", "= (Amount)*2", "=Amount *Units*2", "=Amount*Units *2",
        "=Amount* Units*2", "=(Amount +Units)*2", "=(Amount+ Units)*2", "=( Amount+Units)*2", "=(Amount+Units )*2",
        "=ROUND(Amount ,1)", "=ROUND( Amount,1)", "=ROUND(Amount,1 )", "=ROUND (Amount,1)", "=-Amount * 2",
        "=Amount*-  Units", "=   Amount*Units", "= Amount * Units", "= -Amount", "=2* Amount",
        "=MAX( Amount , Units )", "=IF( Amount>1 ,1,0)", "=Amount+ (Units)", "=(Amount)+ Units",
        "=Amount  +  Units  +  2", "= Amount", "=Amount+Units+ 2", "=Amount+ Units+2", "=(Amount)*(Units) ",
        "=MAX(Amount,  Units)", "=MAX(Amount,Units,  2)", "= MAX(Amount,Units)*2", "=2 * (Amount + Units)",
        "=Amount*(2 + Units)", "=-(Amount + Units)"), 1)))),
    # Subtotals by Sum, captions of the source's fields, a formula without its =, and a cache two tables share.
    "sum_subtotals": (TWO, "", lines('pt.PivotFields("Region").Subtotals(2) = True', RATIO)),
    "source_caption": ("Region", "Product", lines(probe("caption", 'pt.PivotFields("Amount").Caption = "Revenue"'),
                                                  calc("Twice", "=Revenue*2"), calc("Thrice", "=Amount*3"),
                                                  probe("region", 'pt.PivotFields("Region").Caption = "Area"'))),
    "source_name": ("Region", "Product", lines(probe("name", 'pt.PivotFields("Amount").Name = "Revenue"'),
                                               calc("Twice", "=Revenue*2"),
                                               probe("region", 'pt.PivotFields("Region").Name = "Area"'))),
    "no_equals": ("Region", "Product", calc("Twice", "Amount*2")),
    "shared_cache": ("Region", "Product", lines("Set extra = wb.Worksheets.Add(After:=ws)", 'extra.Name = name & "_2"',
                                                'Set other = pt.PivotCache.CreatePivotTable(extra.Range("A3"), name & '
                                                '"_2")', 'other.PivotFields("Region").Orientation = xlRowField',
                                                MARGIN, probe("other", 'other.AddDataField other.PivotFields("Margin")'),
                                                note("other_fields", "other.CalculatedFields.Count"))),
    "caption_lookup": ("Region", "Product", lines(MARGIN, probe("caption", 'pt.PivotFields("Margin").Caption = "Profit"'),
                                                  note("by_caption", 'pt.PivotFields("Profit").Formula'),
                                                  calc("Double", "=Profit*2", data=False),
                                                  calc("Triple", "=Margin*3", data=False),
                                                  note("data_type", 'pt.PivotFields("Region").DataType & "/" & '
                                                                    'pt.PivotFields("Amount").DataType'))),
}

BUILD = HELPER + '''
Public Function Build() As String
Dim wb As Object, data As Object, out As String, r As Long, rows As Variant, parts As Variant
Dim ws As Object
Application.DisplayAlerts = False
Set wb = Workbooks.Add(xlWBATWorksheet)
{source}{layouts}
wb.Worksheets(1).Activate
wb.SaveAs "{folder}\\pivot_calculated_fields.xlsx", 51
wb.Close False
Set wb = Workbooks.Open("{folder}\\pivot_calculated_fields.xlsx")
out = out & "@@"
For Each ws In wb.Worksheets
    If ws.PivotTables.Count > 0 Then out = out & ws.Name & "^" & Dump(ws.PivotTables(1)) & "|"
Next
wb.Close False
Build = out
End Function
'''

LAYOUT = 'out = out & Layout(wb, "{name}", "{row}", "{column}", {step})\n'
REVISIONS = re.compile(r' xmlns:mc="[^"]*"| mc:Ignorable="[^"]*"| xmlns:xr="[^"]*"| xr:uid="[^"]*"| cacheId="\d+"'
                       r'| refreshedDate="[^"]*"| refreshedBy="[^"]*"| r:id="[^"]*"')


def tables(path: Path) -> dict[str, str]:
    """Each pivot table's definition and its cache's, as the file keeps them, by the table's name."""
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
            fields = re.search(r"<cacheFields\b.*?</cacheFields>", cache, re.DOTALL)
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
    strip_save_path(OUT / "pivot_calculated_fields.xlsx")
    dumps = dict(one.split("^", 1) for one in session.split("|") if "^" in one)
    again = dict(one.split("^", 1) for one in reopened.split("|") if "^" in one)
    held = tables(OUT / "pivot_calculated_fields.xlsx")
    record = {"build": BUILD, "source": SOURCE, "layout": LAYOUT, "steps": steps,
              "layouts": [{"name": name, "row": row, "column": column, "step": step,
                           "dump": dumps.get(name, ""), "reopened": again.get(name, ""), "file": held.get(name, "")}
                          for name, (row, column, step) in LAYOUTS.items()],
              # The tables a layout makes beside its own, by name.
              "others": {name: text for name, text in held.items() if name not in LAYOUTS}}
    (OUT / "pivot_calculated_fields.json").write_text(json.dumps(record, indent=1, ensure_ascii=False) + "\n",
                                                      encoding="utf-8")
    for layout in record["layouts"]:
        print("==", layout["name"])
        print("   session ", layout["dump"])
        if layout["reopened"] != layout["dump"]:
            print("   reopened", layout["reopened"])


if __name__ == "__main__":
    main()
