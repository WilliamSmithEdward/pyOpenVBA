"""PivotTable.ShowPages, a sheet for each item of a page field, as Excel makes, names and saves them.

A macro builds a workbook for each scenario: a source list of regions,
products and units, and a pivot table on a sheet of its own with a page
field, and calls ShowPages on it: by the field's name in any case or the
field itself, with no field beside one page field, two or none, with items
hidden or one chosen, on a field that is not a page field, on a table placed
lower down, with its values formatted or a column field, with another
workbook active, and with items that are numbers, dates or blank, whose
names are long, hold apostrophes, no sheet can take, or another sheet has in
any case. A name no field has stops Excel itself, and is left out. One
scenario makes tables with no name given, beside names
taken on the same sheet and on another, to show the numbers Excel gives
them, which run on through the session. The dump lists every sheet in
order, and for each pivot table on it its name, place, page fields' choices
and cells. Each workbook is saved, its pivot table parts kept, and opened and
dumped again. A session of its own logs the events an application's handlers
see as ShowPages makes three sheets.

    python scripts/measure_pivot_show_pages.py

writes tests/fixtures/pivot_show_pages/: each workbook as Excel saved it,
and pivot_show_pages.json, which tests/test_excel_pivot_show_pages.py
replays.
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

OUT = ROOT / "tests" / "fixtures" / "pivot_show_pages"

HELPER = '''Private Function Cells(pt As Object) As String
    Dim out As String, c As Object
    On Error Resume Next
    For Each c In pt.TableRange2.Cells
        out = out & CStr(c.Value2) & ","
    Next
    Cells = out
End Function

Private Function Pages(pt As Object) As String
    Dim out As String, f As Object
    On Error Resume Next
    For Each f In pt.PageFields
        out = out & f.Name & "=" & f.CurrentPage.Name & ";"
    Next
    Pages = out
End Function

Private Function Dump(wb As Object) As String
    Dim out As String, ws As Object, pt As Object
    On Error Resume Next
    For Each ws In wb.Worksheets
        out = out & "[" & ws.Name
        For Each pt In ws.PivotTables
            out = out & " {" & pt.Name & " @" & pt.TableRange2.Address(False, False) & " pages=" & Pages(pt) & _
                " cells=" & Cells(pt) & "}"
        Next
        out = out & "]"
    Next
    Dump = out
End Function

Private Function Made(names As Variant, at As String) As Object
    Dim wb As Object, data As Object, report As Object, pt As Object, r As Long
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set data = wb.Worksheets(1)
    data.Name = "data"
    data.Range("A1:C1").Value = Array("Region", "Product", "Units")
    For r = 0 To 7
        data.Cells(r + 2, 1).Value = names(r Mod (UBound(names) + 1))
        data.Cells(r + 2, 2).Value = Array("Apple", "Banana")(r Mod 2)
        data.Cells(r + 2, 3).Value = r + 1
    Next
    Set report = wb.Worksheets.Add(After:=data)
    report.Name = "report"
    Set pt = wb.PivotCaches.Create(xlDatabase, "data!R1C1:R9C3").CreatePivotTable(report.Range(at), "Report")
    pt.PivotFields("Region").Orientation = xlPageField
    pt.PivotFields("Product").Orientation = xlRowField
    pt.AddDataField pt.PivotFields("Units"), "Sum of Units", xlSum
    Set Made = wb
End Function

Private Function Placed(wb As Object, ws As Object, at As String, name As String) As String
    Dim pt As Object
    On Error Resume Next
    Err.Clear
    If name = "" Then
        Set pt = wb.PivotCaches.Create(xlDatabase, "data!R1C1:R9C3").CreatePivotTable(ws.Range(at))
    Else
        Set pt = wb.PivotCaches.Create(xlDatabase, "data!R1C1:R9C3").CreatePivotTable(ws.Range(at), name)
    End If
    If Err.Number <> 0 Then
        Placed = "E" & Err.Number
    Else
        Placed = pt.Name
    End If
End Function

Private Function Shown(pt As Object) As String
    Dim v As Variant
    v = pt.ShowPages("Region")
    Shown = "returned=" & TypeName(v) & ":" & CStr(v) & " "
End Function

Private Function Numbered(wb As Object) As String
    Dim here As Object, there As Object, first As String, n As Long, out As String
    Set here = wb.Worksheets("report")
    Set there = wb.Worksheets("data")
    first = Placed(wb, here, "H3", "")
    n = CLng(Mid(first, 11))
    out = "names=" & first
    out = out & "," & Placed(wb, here, "L3", "PivotTable" & (n + 1))
    out = out & "," & Placed(wb, here, "P3", "")
    out = out & "," & Placed(wb, there, "H3", "PivotTable" & (n + 3))
    out = out & "," & Placed(wb, here, "T3", "")
    out = out & "," & Placed(wb, here, "X3", "pivottable" & (n + 4))
    out = out & "," & Placed(wb, here, "AB3", "")
    out = out & "," & Placed(wb, here, "AF3", "PIVOTTABLE" & n)
    out = out & "," & Placed(wb, there, "L3", "Report")
    Numbered = out & " "
End Function
'''

REGIONS = 'Array("East", "North", "South", "West")'


def taken(*names: str) -> str:
    """Add a sheet of each name after the others, then show the report again."""
    return "".join(f'wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count)).Name = "{name}"\n'
                   for name in names) + 'wb.Worksheets("report").Activate\n'


#: name -> (the regions the source names, what the scenario does before and with ShowPages).
SCENARIOS: dict[str, tuple[str, str]] = {
    "by_name": (REGIONS, 'pt.ShowPages "Region"'),
    "by_field": (REGIONS, 'pt.ShowPages pt.PivotFields("Region")'),
    "no_field": (REGIONS, "pt.ShowPages"),
    "returned": (REGIONS, "out = Shown(pt)"),
    "chosen": (REGIONS, 'pt.PivotFields("Region").CurrentPage = "North"\npt.ShowPages "Region"'),
    "hidden": (REGIONS, 'pt.PivotFields("Region").EnableMultiplePageItems = True\n'
                        'pt.PivotFields("Region").PivotItems("East").Visible = False\npt.ShowPages "Region"'),
    "single_visible": (REGIONS, 'pt.PivotFields("Region").EnableMultiplePageItems = True\n'
                                'pt.PivotFields("Region").PivotItems("East").Visible = False\n'
                                'pt.PivotFields("Region").PivotItems("North").Visible = False\n'
                                'pt.PivotFields("Region").PivotItems("West").Visible = False\npt.ShowPages "Region"'),
    "two_pages": (REGIONS, 'pt.PivotFields("Product").Orientation = xlPageField\npt.ShowPages "Product"'),
    "not_page": (REGIONS, 'pt.ShowPages "Product"'),
    "anchored": (REGIONS, 'pt.ShowPages "Region"'),
    "odd_names": ('Array("A/B", "Q?", "' + "L" * 35 + '", "History")', 'pt.ShowPages "Region"'),
    "bad_chars": ('Array("a:b", "[x]", "c*", "Plain")', 'pt.ShowPages "Region"'),
    "lengths": ('Array("' + "K" * 31 + '", "' + "J" * 32 + '", "East", "West")', 'pt.ShowPages "Region"'),
    "taken_name": (REGIONS, taken("East") + 'pt.ShowPages "Region"'),
    "taken_twice": (REGIONS, taken("East", "East (2)") + 'pt.ShowPages "Region"'),
    "taken_case": (REGIONS, taken("north") + 'pt.ShowPages "Region"'),
    "taken_long": ('Array("East", "' + "P" * 29 + '", "South", "' + "M" * 31 + '")',
                   taken("P" * 29, "M" * 31) + 'pt.ShowPages "Region"'),
    "item_is_report": ('Array("East", "North", "report", "West")', 'pt.ShowPages "Region"'),
    "numbers": ('Array(10, 20, 30, 40)', 'pt.ShowPages "Region"'),
    "decimals": ('Array(1.5, 2.25, 10, 0.125)', 'pt.ShowPages "Region"'),
    "dates": ('Array(DateSerial(2025, 1, 2), DateSerial(2025, 1, 3), DateSerial(2025, 2, 1), DateSerial(2026, 1, 1))',
              'pt.ShowPages "Region"'),
    "blank": ('Array("East", "", "South", "West")', 'pt.ShowPages "Region"'),
    "active_elsewhere": (REGIONS, 'wb.Worksheets("data").Activate\npt.ShowPages "Region"'),
    "default_names": (REGIONS, "out = Numbered(wb)"),
    "lengths_cut": (f'Array("{"A" * 25}", "{"B" * 26}", "{"C" * 27}", "{"D" * 30}")', 'pt.ShowPages "Region"'),
    "taken_cut": (f'Array("{"E" * 26}", "{"F" * 27}", "Gap", "West")',
                  taken("E" * 26, "F" * 26) + 'pt.ShowPages "Region"'),
    "apostrophes": ('Array("\'\'Lead", "Trail\'", "Mid\'dle", "East")', 'pt.ShowPages "Region"'),
    "two_pages_no_field": (REGIONS, 'pt.PivotFields("Product").Orientation = xlPageField\npt.ShowPages'),
    "no_page_field": (REGIONS, 'pt.PivotFields("Region").Orientation = xlRowField\npt.ShowPages'),
    "lower_name": (REGIONS, 'pt.ShowPages "region"'),
    "other_book_active": (REGIONS, 'Workbooks.Add\npt.ShowPages "Region"\n'
                                   'out = "activebook=" & (ActiveWorkbook Is wb) & " "\n'
                                   "Workbooks(Workbooks.Count).Close False"),
    "formatted": (REGIONS, 'pt.DataFields(1).NumberFormat = "0.00"\npt.ShowPages "Region"'),
    "column_field": (REGIONS, 'pt.PivotFields("Product").Orientation = xlColumnField\npt.ShowPages "Region"'),
    "numbers_no_rows": ('Array(10, 20, 30, 40)',
                        'pt.PivotFields("Product").Orientation = xlColumnField\npt.ShowPages "Region"'),
    "numbers_tabular": ('Array(10, 20, 30, 40)', 'pt.RowAxisLayout xlTabularRow\npt.ShowPages "Region"'),
}

#: Where a scenario puts its table, if not at A3.
ANCHORS = {"anchored": "D10"}

BUILD = HELPER + '''
Private Function Scenario(name As String, names As Variant, at As String, step As Long, folder As String) As String
    Dim wb As Object, pt As Object, out As String
    Set wb = Made(names, at)
    Set pt = wb.Worksheets("report").PivotTables("Report")
    wb.Worksheets("report").Activate
    On Error Resume Next
    Err.Clear
    Select Case step
{steps}
    End Select
    If Err.Number <> 0 Then out = "E" & Err.Number & " " & out
    On Error GoTo 0
    out = out & "active=" & wb.ActiveSheet.Name & " " & Dump(wb)
    wb.SaveAs folder & "\\" & name & ".xlsx", 51
    wb.Close False
    Set wb = Workbooks.Open(folder & "\\" & name & ".xlsx")
    out = out & "@@" & Dump(wb)
    wb.Close False
    Scenario = name & "^" & out & "|"
End Function

Public Function Build() As String
Dim out As String
Application.DisplayAlerts = False
{scenarios}
Build = out
End Function
'''

SCENARIO = 'out = out & Scenario("{name}", {names}, "{at}", {step}, "{folder}")\n'

#: An application's handlers, which log the events ShowPages raises.
SINK = '''Public WithEvents App As Application
Private Sub App_WorkbookNewSheet(ByVal Wb As Workbook, ByVal Sh As Object)
Log = Log & "new:" & Sh.Name & "(" & Sh.PivotTables.Count & "," & Sh.Range("A1").Value & "," & Sh.Range("B4").Value & "," & Sh.Index & "," & ActiveSheet.Name & "," & Sh.Parent.Worksheets.Count & ")|"
End Sub
Private Sub App_SheetActivate(ByVal Sh As Object)
Log = Log & "activate:" & Sh.Name & "(" & ActiveSheet.Name & ")|"
End Sub
Private Sub App_SheetDeactivate(ByVal Sh As Object)
Log = Log & "deactivate:" & Sh.Name & "|"
End Sub
Private Sub App_SheetPivotTableUpdate(ByVal Sh As Object, ByVal Target As PivotTable)
Log = Log & "update:" & Sh.Name & "/" & Target.Name & "(" & ActiveSheet.Name & ")|"
End Sub
Private Sub App_SheetChange(ByVal Sh As Object, ByVal Target As Range)
Log = Log & "change:" & Sh.Name & "!" & Target.Address(False, False) & "(" & Target.Cells(1, 1).Value & ")|"
End Sub
Private Sub App_SheetCalculate(ByVal Sh As Object)
Log = Log & "calculate:" & Sh.Name & "|"
End Sub
Private Sub App_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
Log = Log & "select:" & Sh.Name & "!" & Target.Address(False, False) & "|"
End Sub
'''

WATCH = HELPER + '''
Public Function Watch() As String
    Dim wb As Object, pt As Object, sink As New ShowSink, out As String
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Set wb = Made(Array("East", "North", "South", "West"), "A3")
    Set pt = wb.Worksheets("report").PivotTables("Report")
    pt.PivotFields("Region").CurrentPage = "North"
    wb.Worksheets("report").Activate
    Set sink.App = Application
    Application.EnableEvents = True
    Log = ""
    pt.ShowPages "Region"
    out = Log
    Application.EnableEvents = False
    out = out & "active=" & ActiveSheet.Name & " selection=" & Selection.Address(False, False)
    wb.Close False
    Watch = out
End Function
'''


def install(code: str, name: str) -> str:
    """A macro adding a class module of ``code`` to the workbook the harness runs in."""
    lines = ["Function Install() As String", "Dim text As String, part As Object",
             "Set part = ActiveWorkbook.VBProject.VBComponents.Add(2)", f'part.Name = "{name}"', 'text = ""']
    lines += ['text = text & "' + line.replace('"', '""') + '" & vbLf' for line in code.splitlines()]
    return "\n".join([*lines, "part.CodeModule.AddFromString text", 'Install = "ok"', "End Function"])


def events() -> dict[str, str]:
    """The events ShowPages raises, in a session of their own, so the table names count from 1."""
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        excel.add_module("EventLog", "Public Log As String\n")
        installed = excel.run_vba(install(SINK, "ShowSink"), "Install", timeout=60.0)
        assert installed.ok, f"{installed.outcome}: {installed.message}"
        result = excel.run_vba(WATCH, "Watch", timeout=300.0)
        assert result.ok, f"{result.outcome}: {result.message} {result.error}"
    return {"sink": SINK, "watch": WATCH, "log": str(result.value)}
REVISIONS = re.compile(r' xmlns:mc="[^"]*"| mc:Ignorable="[^"]*"| xmlns:xr="[^"]*"| xr:uid="[^"]*"| cacheId="\d+"'
                       r'| refreshedDate="[^"]*"| refreshedBy="[^"]*"| r:id="[^"]*"')


def tables(path: Path) -> dict[str, str]:
    """Each pivot table's definition, by its sheet's name and the table's, as the file keeps it."""
    with zipfile.ZipFile(path) as package:
        workbook = package.read("xl/workbook.xml").decode("utf-8")
        rels = package.read("xl/_rels/workbook.xml.rels").decode("utf-8")
        targets = dict(re.findall(r'Id="([^"]+)"[^>]*Target="([^"]+)"', rels))
        targets.update({key: value for value, key in re.findall(r'Target="([^"]+)"[^>]*Id="([^"]+)"', rels)})
        out: dict[str, str] = {}
        for sheet, relation in re.findall(r'<sheet\b[^>]*name="([^"]*)"[^>]*r:id="([^"]+)"', workbook):
            part = posixpath.normpath(posixpath.join("xl", targets[relation]))
            sheet_rels = posixpath.join(posixpath.dirname(part), "_rels", posixpath.basename(part) + ".rels")
            if sheet_rels not in package.namelist():
                continue
            for target in re.findall(r'Target="([^"]*pivotTable\d+\.xml)"', package.read(sheet_rels).decode()):
                table = posixpath.normpath(posixpath.join(posixpath.dirname(part), target))
                text = re.sub(r"^<\?xml[^>]*\?>\s*", "", package.read(table).decode("utf-8"))
                name = re.search(r'<pivotTableDefinition\b[^>]*?\bname="([^"]*)"', text)
                out[f"{sheet}/{name.group(1) if name else table}"] = REVISIONS.sub("", text)
    return out


def main() -> None:
    wanted = set(sys.argv[1:])
    if not wanted and OUT.exists():
        shutil.rmtree(OUT)
    OUT.mkdir(parents=True, exist_ok=True)
    steps = "".join(f"    Case {index}\n" + "".join(f"        {line}\n" for line in step.splitlines())
                    for index, (_names, step) in enumerate(SCENARIOS.values(), 1))
    body = "".join(SCENARIO.replace("{name}", name).replace("{names}", names)
                   .replace("{at}", ANCHORS.get(name, "A3")).replace("{step}", str(index))
                   .replace("{folder}", str(OUT))
                   for index, (name, (names, _step)) in enumerate(SCENARIOS.items(), 1)
                   if not wanted or name in wanted)
    code = BUILD.replace("{steps}", steps).replace("{scenarios}", body)
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(code, "Build", timeout=600.0)
        assert result.ok, f"{result.outcome}: {result.message} {result.error}"
    found = dict(one.split("^", 1) for one in str(result.value).split("|") if "^" in one)
    record: dict[str, object] = {"build": BUILD, "steps": steps, "scenario": SCENARIO, "events": events(),
                                 "scenarios": []}
    scenarios: list[dict[str, object]] = []
    record["scenarios"] = scenarios
    for name, (names, step) in SCENARIOS.items():
        if wanted and name not in wanted:
            continue
        strip_save_path(OUT / f"{name}.xlsx")
        session, reopened = found.get(name, "@@").split("@@", 1)
        scenarios.append({"name": name, "names": names, "at": ANCHORS.get(name, "A3"), "step": step,
                          "dump": session, "reopened": reopened, "files": tables(OUT / f"{name}.xlsx")})
        print("==", name)
        print("   session ", session[:1200])
        if reopened != session.split(" ", 1)[-1]:
            print("   reopened", reopened[:1200])
    (OUT / "pivot_show_pages.json").write_text(json.dumps(record, indent=1, ensure_ascii=False) + "\n",
                                               encoding="utf-8")


if __name__ == "__main__":
    main()
