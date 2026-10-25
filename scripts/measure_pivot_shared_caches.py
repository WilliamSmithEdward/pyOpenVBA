"""Pivot tables sharing a cache across sheets: refreshing, RefreshAll and the cache's settings, in live Excel.

A macro builds a workbook for each scenario: a source list of regions,
products and amounts, a cache, two tables made from it on two sheets and a
third beside the first, and for some one or two caches more over other
ranges, each with a table of its own. It changes the source and refreshes
one table, a cache or the workbook, a cache first for some, or sets a
setting of the cache on one table, a style flag waiting on another and the
like. The dump lists every sheet and for
each table its name, its cache's index, its place, its cells and the widths
of the sheet's first eight columns. Each workbook is saved, its pivot parts
kept, and opened and dumped again. A session of its own logs the events an
application's handlers see as a refresh lays out tables on two sheets.

    python scripts/measure_pivot_shared_caches.py [scenario ...]

writes tests/fixtures/pivot_shared_caches/: each workbook as Excel saved it
and pivot_shared_caches.json, which tests/test_excel_pivot_shared_caches.py
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

OUT = ROOT / "tests" / "fixtures" / "pivot_shared_caches"

HELPER = '''Private Function Cells(pt As Object) As String
    Dim out As String, c As Object
    On Error Resume Next
    For Each c In pt.TableRange2.Cells
        out = out & CStr(c.Value2) & ","
    Next
    Cells = out
End Function

Private Function Widths(ws As Object) As String
    Dim c As Long, out As String
    For c = 1 To 8
        out = out & Round(ws.Columns(c).Width * 4 / 3) & ","
    Next
    Widths = out
End Function

Private Function Dump(wb As Object) As String
    Dim out As String, ws As Object, pt As Object
    On Error Resume Next
    For Each ws In wb.Worksheets
        out = out & "[" & ws.Name & " w=" & Widths(ws)
        For Each pt In ws.PivotTables
            out = out & " {" & pt.Name & " cache=" & pt.CacheIndex & " @" & pt.TableRange2.Address(False, False) & _
                " cells=" & Cells(pt) & "}"
        Next
        out = out & "]"
    Next
    out = out & " caches=" & wb.PivotCaches.Count
    Dump = out
End Function

Private Function Made(extra As Long) As Object
    Dim wb As Object, data As Object, a As Object, b As Object, c As Object, d As Object, cache As Object
    Dim other As Object, third As Object, r As Long
    Dim t1 As Object, t2 As Object, t3 As Object, t4 As Object, t5 As Object
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set data = wb.Worksheets(1)
    data.Name = "data"
    data.Range("A1:C1").Value = Array("Region", "Product", "Amount")
    For r = 0 To 7
        data.Cells(r + 2, 1).Value = Array("East", "North", "South", "West")(r Mod 4)
        data.Cells(r + 2, 2).Value = Array("Apple", "Banana")(r Mod 2)
        data.Cells(r + 2, 3).Value = (r + 1) * 10
    Next
    data.Range("E1:F1").Value = Array("Channel", "Units")
    For r = 0 To 3
        data.Cells(r + 2, 5).Value = Array("Online", "Store")(r Mod 2)
        data.Cells(r + 2, 6).Value = r + 1
    Next
    data.Range("H1:I1").Value = Array("Kind", "Count")
    For r = 0 To 3
        data.Cells(r + 2, 8).Value = Array("New", "Used")(r Mod 2)
        data.Cells(r + 2, 9).Value = (r + 1) * 3
    Next
    Set a = wb.Worksheets.Add(After:=data)
    a.Name = "one"
    Set b = wb.Worksheets.Add(After:=a)
    b.Name = "two"
    Set cache = wb.PivotCaches.Create(xlDatabase, "data!R1C1:R9C3")
    Set t1 = cache.CreatePivotTable(a.Range("A3"), "First")
    t1.PivotFields("Region").Orientation = xlRowField
    t1.AddDataField t1.PivotFields("Amount"), "Sum of Amount", xlSum
    Set t2 = cache.CreatePivotTable(b.Range("A3"), "Second")
    t2.PivotFields("Product").Orientation = xlRowField
    t2.AddDataField t2.PivotFields("Amount"), "Total", xlSum
    Set t3 = cache.CreatePivotTable(a.Range("F3"), "Third")
    t3.PivotFields("Product").Orientation = xlColumnField
    t3.AddDataField t3.PivotFields("Amount"), "Sum of Amount", xlSum
    If extra > 0 Then
        Set c = wb.Worksheets.Add(After:=b)
        c.Name = "three"
        Set other = wb.PivotCaches.Create(xlDatabase, "data!R1C5:R5C6")
        Set t4 = other.CreatePivotTable(c.Range("A3"), "Fourth")
        t4.PivotFields("Channel").Orientation = xlRowField
        t4.AddDataField t4.PivotFields("Units"), "Sum of Units", xlSum
    End If
    If extra > 1 Then
        Set d = wb.Worksheets.Add(After:=c)
        d.Name = "four"
        Set third = wb.PivotCaches.Create(xlDatabase, "data!R1C8:R5C9")
        Set t5 = third.CreatePivotTable(d.Range("A3"), "Fifth")
        t5.PivotFields("Kind").Orientation = xlRowField
        t5.AddDataField t5.PivotFields("Count"), "Sum of Count", xlSum
    End If
    Set Made = wb
End Function

Private Function Indexes(wb As Object) As String
    Dim out As String, ws As Object, pt As Object
    For Each ws In wb.Worksheets
        For Each pt In ws.PivotTables
            out = out & pt.Name & "=" & pt.CacheIndex & ","
        Next
    Next
    Indexes = out
End Function
'''

#: The edit the source gets before a refresh: a longer region name and a new amount.
EDIT = 'wb.Worksheets("data").Range("A2").Value = "A region name longer than the rest"\n' \
       'wb.Worksheets("data").Range("C3").Value = 1000\n'

#: name -> (how many caches are made beside the shared one, what the scenario does).
SCENARIOS: dict[str, tuple[int, str]] = {
    "made": (1, ""),
    "refresh_table": (0, EDIT + 'wb.Worksheets("two").PivotTables("Second").RefreshTable'),
    "refresh_cache": (0, EDIT + 'wb.Worksheets("one").PivotTables("First").PivotCache.Refresh'),
    "refresh_by_collection": (1, EDIT + "wb.PivotCaches(1).Refresh"),
    "refresh_all": (1, EDIT + 'wb.Worksheets("data").Range("F2").Value = 50\nwb.RefreshAll'),
    "refresh_second_cache": (1, 'wb.Worksheets("data").Range("F2").Value = 50\n'
                                   'wb.Worksheets("three").PivotTables("Fourth").PivotCache.Refresh'),
    "refresh_unchanged": (0, 'wb.Worksheets("one").PivotTables("Third").RefreshTable'),
    "flags_then_refresh": (0, 'wb.Worksheets("one").PivotTables("Third").ShowTableStyleColumnHeaders = False\n'
                                  'wb.Worksheets("two").PivotTables("Second").RefreshTable'),
    "flags_then_update": (0, 'wb.Worksheets("one").PivotTables("Third").ShowTableStyleColumnHeaders = False\n'
                                 'wb.Worksheets("two").PivotTables("Second").Update'),
    "save_data": (0, 'wb.Worksheets("one").PivotTables("Third").ShowTableStyleColumnHeaders = False\n'
                         'wb.Worksheets("two").PivotTables("Second").SaveData = False\n'
                         'out = "first=" & wb.Worksheets("one").PivotTables("First").SaveData & " "'),
    "calculated_field": (0, 'wb.Worksheets("one").PivotTables("Third").ShowTableStyleColumnHeaders = False\n'
                                'wb.Worksheets("two").PivotTables("Second").CalculatedFields.Add "Double", "=Amount*2"\n'
                                'out = "first=" & wb.Worksheets("one").PivotTables("First").CalculatedFields.Count & " "'),
    "refresh_disabled": (0, 'wb.Worksheets("two").PivotTables("Second").PivotCache.EnableRefresh = False\n'
                                'wb.Worksheets("one").PivotTables("First").RefreshTable'),
    "cache_index": (1, 'out = "before=" & wb.Worksheets("three").PivotTables("Fourth").CacheIndex & ","\n'
                          'wb.Worksheets("one").PivotTables("First").PivotCache.Refresh\n'
                          'out = out & "after=" & wb.Worksheets("three").PivotTables("Fourth").CacheIndex & ","\n'
                          'out = out & wb.Worksheets("two").PivotTables("Second").CacheIndex & " "'),
    "new_item": (0, 'wb.Worksheets("data").Range("A9").Value = "Central"\n'
                        'wb.Worksheets("one").PivotTables("First").RefreshTable'),
    "refresh_all_disabled": (1, EDIT + 'wb.Worksheets("data").Range("F2").Value = 50\n'
                                'wb.Worksheets("three").PivotTables("Fourth").PivotCache.EnableRefresh = False\n'
                                "wb.RefreshAll"),
    "refresh_all_three": (2, 'out = "before=" & Indexes(wb) & " "\nwb.RefreshAll\n'
                             'out = out & "after=" & Indexes(wb) & " "'),
    "refresh_all_reordered": (2, 'wb.Worksheets("three").PivotTables("Fourth").PivotCache.Refresh\n'
                                 'out = "before=" & Indexes(wb) & " "\nwb.RefreshAll\n'
                                 'out = out & "after=" & Indexes(wb) & " "'),
    "item_gone": (0, 'wb.Worksheets("data").Range("B3").Value = "Apple"\n'
                         'wb.Worksheets("data").Range("B5").Value = "Apple"\n'
                         'wb.Worksheets("data").Range("B7").Value = "Apple"\n'
                         'wb.Worksheets("data").Range("B9").Value = "Apple"\n'
                         'wb.Worksheets("one").PivotTables("First").RefreshTable'),
}

BUILD = HELPER + '''
Private Function Scenario(name As String, extra As Long, step As Long, folder As String) As String
    Dim wb As Object, out As String
    Set wb = Made(extra)
    wb.Worksheets("one").Activate
    On Error Resume Next
    Err.Clear
    Select Case step
{steps}
    End Select
    If Err.Number <> 0 Then out = "E" & Err.Number & " " & out
    On Error GoTo 0
    out = out & Dump(wb)
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

SCENARIO = 'out = out & Scenario("{name}", {extra}, {step}, "{folder}")\n'

#: An application's handlers, which log the events a refresh over two sheets raises.
SINK = '''Public WithEvents App As Application
Private Sub App_SheetPivotTableUpdate(ByVal Sh As Object, ByVal Target As PivotTable)
Log = Log & "update:" & Sh.Name & "/" & Target.Name & "|"
End Sub
Private Sub App_SheetChange(ByVal Sh As Object, ByVal Target As Range)
Log = Log & "change:" & Sh.Name & "!" & Target.Address(False, False) & "|"
End Sub
Private Sub App_SheetCalculate(ByVal Sh As Object)
Log = Log & "calculate:" & Sh.Name & "|"
End Sub
Private Sub App_SheetActivate(ByVal Sh As Object)
Log = Log & "activate:" & Sh.Name & "|"
End Sub
'''

WATCH = HELPER + '''
Public Function Watch(how As Long) As String
    Dim wb As Object, sink As New SharedSink, out As String
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Set wb = Made(2)
    wb.Worksheets("one").Activate
    wb.Worksheets("data").Range("C3").Value = 1000
    If how = 3 Then wb.Worksheets("three").PivotTables("Fourth").PivotCache.Refresh
    Set sink.App = Application
    Application.EnableEvents = True
    Log = ""
    Select Case how
        Case 1
            wb.Worksheets("two").PivotTables("Second").RefreshTable
        Case 2, 3
            wb.RefreshAll
    End Select
    out = Log
    Application.EnableEvents = False
    out = out & "active=" & ActiveSheet.Name
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


def events() -> dict[str, object]:
    """The events a refresh of a cache over two sheets raises, and RefreshAll's, in a session of their own."""
    logs: dict[str, str] = {}
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        excel.add_module("EventLog", "Public Log As String\n")
        installed = excel.run_vba(install(SINK, "SharedSink"), "Install", timeout=60.0)
        assert installed.ok, f"{installed.outcome}: {installed.message}"
        for how, label in ((1, "refresh_table"), (2, "refresh_all"), (3, "refresh_all_reordered")):
            result = excel.run_vba(WATCH, "Watch", (how,), timeout=300.0)
            assert result.ok, f"{result.outcome}: {result.message} {result.error}"
            logs[label] = str(result.value)
    return {"sink": SINK, "watch": WATCH, "logs": logs}


def parts(path: Path) -> dict[str, str]:
    """The pivot parts as the file keeps them, by part name, stamps left out."""
    stamps = re.compile(r' xmlns:mc="[^"]*"| mc:Ignorable="[^"]*"| xmlns:xr="[^"]*"| xr:uid="[^"]*"'
                        r'| refreshedDate="[^"]*"| refreshedBy="[^"]*"| r:id="[^"]*"')
    with zipfile.ZipFile(path) as package:
        return {name: stamps.sub("", re.sub(r"^<\?xml[^>]*\?>\s*", "", package.read(name).decode("utf-8")))
                for name in sorted(package.namelist()) if name.startswith("xl/pivot")
                or name == "xl/workbook.xml"}


def main() -> None:
    wanted = set(sys.argv[1:])
    if not wanted and OUT.exists():
        shutil.rmtree(OUT)
    OUT.mkdir(parents=True, exist_ok=True)
    steps = "".join(f"    Case {index}\n" + "".join(f"        {line}\n" for line in step.splitlines())
                    for index, (_extra, step) in enumerate(SCENARIOS.values(), 1))
    body = "".join(SCENARIO.replace("{name}", name).replace("{extra}", str(extra))
                   .replace("{step}", str(index)).replace("{folder}", str(OUT))
                   for index, (name, (extra, _step)) in enumerate(SCENARIOS.items(), 1)
                   if not wanted or name in wanted)
    code = BUILD.replace("{steps}", steps).replace("{scenarios}", body)
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(code, "Build", timeout=900.0)
        assert result.ok, f"{result.outcome}: {result.message} {result.error}"
    found = dict(one.split("^", 1) for one in str(result.value).split("|") if "^" in one)
    record: dict[str, object] = {"build": BUILD, "steps": steps, "scenario": SCENARIO, "events": events()}
    scenarios: list[dict[str, object]] = []
    record["scenarios"] = scenarios
    for name, (extra, step) in SCENARIOS.items():
        if wanted and name not in wanted:
            continue
        strip_save_path(OUT / f"{name}.xlsx")
        session, reopened = found.get(name, "@@").split("@@", 1)
        scenarios.append({"name": name, "extra": extra, "step": step, "dump": session, "reopened": reopened,
                          "files": parts(OUT / f"{name}.xlsx")})
        print("==", name)
        print("   session ", session[:900])
    (OUT / "pivot_shared_caches.json").write_text(json.dumps(record, indent=1, ensure_ascii=False) + "\n",
                                                  encoding="utf-8")


if __name__ == "__main__":
    main()
