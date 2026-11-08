"""Charts.Add, and Sheets.Add with Type xlChart: what the new chart sheet plots, where it lands, what the arguments do, and
the parts a save writes for it.

Each case makes a workbook of three worksheets, One, Two and Three, with One active, a small table in its A1:C3 and A1
selected, runs its statements, and records r, a digest of the tabs, and the parts a chart sheet touches.

    python scripts/measure_charts_add.py

writes tests/fixtures/charts_add.json, which tests/test_excel_charts_add.py replays.
"""
from __future__ import annotations

import json
import re
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/charts_add.json'
ADD = 'Set t = wb.Charts.Add: '
# What the new chart plots: the selection, or a single cell's current region, of the active worksheet.
CASES: list[str] = [ADD + 'r = Seen(t)', 'ws.Range("B2").Select: ' + ADD + 'r = Seen(t)', 'ws.Range("A1:B3").Select: ' + ADD + 'r = Seen(t)',
                    'ws.Range("B1:C3").Select: ' + ADD + 'r = Seen(t)', 'ws.Range("E10").Select: ' + ADD + 'r = Seen(t)',
                    'ws.Range("A1:A3,C1:C3").Select: ' + ADD + 'r = Seen(t)', 'ws.Range("A1:C1").Select: ' + ADD + 'r = Seen(t)',
                    'ws.Range("B2:C3").Select: ' + ADD + 'r = Seen(t)', 'ws.Range("A1:C3").Select: ' + ADD + 'r = Seen(t)',
                    'wb.Worksheets("Two").Activate: ' + ADD + 'r = Seen(t)', 'ws.Range("A1:C3").ClearContents: ' + ADD + 'r = Seen(t)',
                    'ws.Range("A4:C4").Value = Array("c", 30, 8): ' + ADD + 'r = Seen(t)', 'ws.Range("C2").Value = "": ' + ADD + 'r = Seen(t)',
                    'ws.Range("A1:C3").Select: wb.Worksheets("Two").Activate: ' + ADD + 'r = Seen(t)',
                    'wb.Worksheets("Two").Range("A1:B2").Value = 5: wb.Worksheets("Two").Activate: ' + ADD + 'r = Seen(t)',
                    'ws.Shapes.AddShape(1, 10, 10, 10, 10).Select: ' + ADD + 'r = Seen(t)',
                    ADD + 'Set u = wb.Charts.Add: r = Seen(u)', 'Set t = wb.Charts.Add(Count:=2): r = Seen(t) & " / " & Seen(wb.Charts(2))',
                    'Set t = wb.Sheets.Add(Type:=xlChart): r = Seen(t)', 'ws.Range("A1:B3").Select: ' + ADD + 'r = t.HasTitle & " " & t.ChartTitle.Text',
                    ADD + 'r = t.SeriesCollection(1).Name & " " & t.SeriesCollection(2).Name', ADD + 'r = t.ChartStyle & " " & t.Legend.Position & " " & t.PlotBy']
# The chart sheet once made.
CASES += [ADD + 'r = TypeName(ActiveChart) & " " & ActiveChart.Name & " " & ActiveSheet.Name & " " & TypeName(t.Parent)', ADD + 'ws.Range("B2").Value = 99: r = t.SeriesCollection(1).Values(1)',
          ADD + 't.ChartType = xlLine: r = t.ChartType', ADD + 't.SetSourceData ws.Range("A1:B3"): r = Seen(t)', ADD + 't.HasTitle = True: t.ChartTitle.Text = "T": r = t.ChartTitle.Text',
          ADD + 't.Delete: r = wb.Sheets.Count & " " & ActiveSheet.Name', ADD + 't.Name = "Mine": r = t.Name', ADD + 'ws.Activate: r = TypeName(ActiveChart) & " " & ActiveSheet.Name',
          ADD + 'r = wb.Windows(1).SelectedSheets.Count & " " & t.Visible', ADD + 'Set u = wb.Charts.Add(After:=t): r = u.Index & " " & u.Name']
# Where it lands, and the arguments.
CASES += ['Set t = wb.Charts.Add(Before:=wb.Sheets(2)): r = t.Index', 'Set t = wb.Charts.Add(After:=wb.Sheets(1)): r = t.Index', 'Set t = wb.Charts.Add(After:=wb.Sheets(2)): r = t.Index',
          'Set t = wb.Charts.Add(After:=wb.Sheets(3)): r = t.Index', 'Set t = wb.Charts.Add(Before:=wb.Sheets(1), After:=wb.Sheets(2)): r = t.Index',
          'Set t = wb.Charts.Add(Count:=0): r = 1', 'Set t = wb.Charts.Add(Count:=-1): r = 1', 'Set t = wb.Charts.Add(Count:="x"): r = 1', 'Set t = wb.Charts.Add(Count:=2.7): r = wb.Charts.Count',
          'Set t = wb.Charts.Add(After:=wb.Sheets(1), Count:=2): r = t.Name & " " & t.Index', 'Set t = wb.Charts.Add(After:=wb.Sheets(3), Count:=2): r = t.Name & " " & t.Index',
          'wb.Worksheets("Three").Activate: ' + ADD + 'r = t.Index', 'Set t = wb.Charts.Add(Before:="Two"): r = 1',
          'wb.Protect Structure:=True: ' + ADD + 'r = 1', 'ws.Protect: ' + ADD + 'r = Seen(t)', 'wb.Worksheets.Add(After:=ws).Name = "Chart1": ws.Activate: ' + ADD + 'r = t.Name',
          ADD + 'Application.DisplayAlerts = False: t.Delete: Set t = wb.Charts.Add: r = t.Name', 'Set t = wb.Charts.Add(Type:=xlChart): r = Seen(t)',
          'Set t = wb.Charts.Add(Type:=xlWorksheet): r = TypeName(t) & " " & t.Name', 'Set t = wb.Sheets.Add(Type:=xlChart, Count:=2): r = t.Name & " " & t.Index',
          'Set t = wb.Sheets.Add(After:=wb.Sheets(3), Type:=xlChart): r = t.Index', 'Set t = wb.Worksheets.Add(Type:=xlChart): r = TypeName(t)']
# A chart sheet made while a chart is active plots what the first did. The rivals: (P1) a copy of the active chart's series,
# (P2) the selection of the worksheet active before the chart sheets, (P3) the first worksheet's. And with a shape selected:
# (A1) the active cell's region, (A2) the range selected before the shape.
TWO = 'wb.Worksheets("Two").Range("A1:B2").Value = 5: '
CASES += [ADD + 't.SetSourceData ws.Range("A1:B3"): Set u = wb.Charts.Add: r = Seen(u)',
          TWO + 'wb.Worksheets("Two").Activate: ' + ADD + 'Set u = wb.Charts.Add: r = Seen(u)',
          TWO + ADD + 't.SetSourceData wb.Worksheets("Two").Range("A1:B2"): Set u = wb.Charts.Add: r = Seen(u)',
          ADD + 't.ChartType = xlLine: Set u = wb.Charts.Add: r = Seen(u)',
          'ws.Range("E10").Select: ' + ADD + 't.SetSourceData ws.Range("A1:C3"): Set u = wb.Charts.Add: r = Seen(u)',
          ADD + 't.HasTitle = True: t.ChartTitle.Text = "T": Set u = wb.Charts.Add: r = u.HasTitle & " " & Seen(u)',
          ADD + 't.SetSourceData ws.Range("A1:B3"): ws.Activate: t.Activate: Set u = wb.Charts.Add: r = Seen(u)',
          'Set co = ws.ChartObjects.Add(100, 100, 200, 100): co.Chart.SetSourceData ws.Range("A1:B3"): co.Activate: ' + ADD + 'r = Seen(t)',
          'Set co = ws.ChartObjects.Add(100, 100, 200, 100): co.Chart.SetSourceData ws.Range("A1:B3"): co.Select: ' + ADD + 'r = Seen(t)',
          'ws.Range("A1:B3").Select: ws.Shapes.AddShape(1, 10, 10, 10, 10).Select: ' + ADD + 'r = Seen(t)',
          'ws.Range("E10").Select: ws.Shapes.AddShape(1, 10, 10, 10, 10).Select: ' + ADD + 'r = Seen(t)',
          ADD + 'ws.Range("A4:C4").Value = Array("c", 30, 8): Set u = wb.Charts.Add: r = Seen(u)',
          'ws.Range("A1:B3").Select: ' + TWO + 'wb.Worksheets("Two").Activate: ' + ADD + 'ws.Activate: Set u = wb.Charts.Add: r = Seen(u)',
          ADD + 't.SeriesCollection(2).Delete: Set u = wb.Charts.Add: r = Seen(u)']
# Held out: the rule from the cases above, the active chart's last source plotted again, on cases written after it. The rivals:
# the kept range follows inserts and cuts as a held range does (K1) or stays at its address (K2); SetSourceData's PlotBy is kept
# (B1) or guessed again (B2); NewSeries leaves the source as it was (N1) or adds to it (N2).
CO = 'Set co = ws.ChartObjects.Add(100, 100, 200, 100): '
CASES += [ADD + 'ws.Rows(1).Insert: Set u = wb.Charts.Add: r = Seen(u)', ADD + 'ws.Columns(1).Insert: Set u = wb.Charts.Add: r = Seen(u)',
          ADD + 't.SetSourceData ws.Range("A1:C3"), xlColumns: Set u = wb.Charts.Add: r = Seen(u)', ADD + 't.SeriesCollection.NewSeries: Set u = wb.Charts.Add: r = Seen(u)',
          'Set t = wb.Charts.Add(Count:=3): r = Seen(wb.Charts(3))',
          CO + 'co.Chart.SetSourceData ws.Range("A1:B3"): co.Activate: Set t = wb.Sheets.Add(Type:=xlChart): r = Seen(t)',
          'ws.Shapes.AddChart2(201, xlColumnClustered).Chart.SetSourceData ws.Range("B1:C3"): ws.ChartObjects(1).Activate: ' + ADD + 'r = Seen(t)',
          CO + 'co.Chart.SetSourceData ws.Range("A1:B3"): Set t = co.Chart.Location(xlLocationAsNewSheet): Set u = wb.Charts.Add: r = Seen(u)',
          ADD + 'ws.Range("A1:C3").ClearContents: Set u = wb.Charts.Add: r = Seen(u)', ADD + 'ws.Activate: ws.Range("E10").Select: t.Activate: Set u = wb.Charts.Add: r = Seen(u)',
          ADD + 'ws.Activate: ws.Range("A1:B3").Select: Set u = wb.Charts.Add: r = Seen(u)', ADD + 't.SetSourceData ws.Range("A1:A3,C1:C3"): Set u = wb.Charts.Add: r = Seen(u)',
          CO + 'co.Activate: ' + ADD + 'r = Seen(t)',
          ADD + 'wb.Worksheets("Two").Activate: wb.Worksheets("Two").Range("B2").Select: t.Activate: Set u = wb.Charts.Add: r = Seen(u)',
          ADD + 'Set u = wb.Charts.Add(Before:=wb.Sheets(4)): r = u.Index & " " & Seen(u)', ADD + 'ws.Range("A1:C3").Cut ws.Range("E5"): Set u = wb.Charts.Add: r = Seen(u)']
# With Before or After a chart active plotted nothing. The rivals: (W1) what is plotted is the anchor's, a worksheet's selection or
# a chart's last source, or (W2) the active sheet's as without them, a chart sheet active giving nothing.
CASES += ['Set t = wb.Charts.Add(Before:=wb.Sheets(2)): r = Seen(t)', 'Set t = wb.Charts.Add(After:=wb.Sheets(1)): r = Seen(t)',
          TWO + 'Set t = wb.Charts.Add(Before:=wb.Worksheets("Two")): r = Seen(t)', TWO + 'Set t = wb.Charts.Add(After:=wb.Worksheets("Two")): r = Seen(t)',
          ADD + 'Set u = wb.Charts.Add(After:=t): r = Seen(u)', ADD + 'Set u = wb.Charts.Add(Before:=ws): r = Seen(u)',
          TWO + ADD + 'Set u = wb.Charts.Add(Before:=wb.Worksheets("Two")): r = Seen(u)', ADD + 'Set u = wb.Charts.Add(Before:=t): r = Seen(u)',
          TWO + 'Set t = wb.Sheets.Add(Type:=xlChart, Before:=wb.Worksheets("Two")): r = Seen(t)',
          TWO + 'Set t = wb.Charts.Add(Before:=wb.Worksheets("Two"), Count:=2): r = Seen(t) & " / " & Seen(wb.Charts(2))',
          'ws.Range("A1:B3").Select: Set t = wb.Charts.Add(Before:=ws): r = Seen(t)',
          TWO + 'wb.Worksheets("Two").Activate: wb.Worksheets("Two").Range("B2").Select: ws.Activate: Set t = wb.Charts.Add(Before:=wb.Worksheets("Two")): r = Seen(t)',
          'ws.Range("A1:B3").Select: ' + TWO + 'wb.Worksheets("Two").Activate: Set t = wb.Charts.Add(Before:=ws): r = Seen(t)',
          ADD + 't.SetSourceData ws.Range("A1:B3"): Set u = wb.Charts.Add(After:=t): r = Seen(u)']
# Held out: the rule from the cases above, a new chart sheet plotting what the tab it goes before has (a worksheet's selection or
# its active chart's last source, a chart sheet's last source, in the orientation that chart was plotted), on cases written after it.
CASES += [TWO + ADD + 'Set u = wb.Charts.Add(After:=wb.Worksheets("One")): r = Seen(u)', ADD + 'ws.Activate: Set u = wb.Charts.Add(Before:=t): r = Seen(u)',
          ADD + 't.SetSourceData ws.Range("A1:B3"): ws.Activate: Set u = wb.Charts.Add(Before:=t): r = Seen(u)',
          CO + 'co.Chart.SetSourceData ws.Range("A1:B3"): co.Activate: Set t = wb.Charts.Add(Before:=ws): r = Seen(t)',
          CO + 'co.Chart.SetSourceData ws.Range("A1:B3"): co.Activate: Set t = wb.Charts.Add(Before:=wb.Worksheets("Two")): r = Seen(t)',
          TWO + 'wb.Worksheets("Two").Activate: wb.Worksheets("Two").Range("A1:A2").Select: ws.Activate: Set t = wb.Charts.Add(After:=ws): r = Seen(t)',
          TWO + 'Set t = wb.Charts.Add(After:=wb.Worksheets("Three")): r = Seen(t)',
          'wb.Worksheets("Three").Range("A1:B2").Value = 7: Set t = wb.Charts.Add(After:=wb.Worksheets("Three")): r = Seen(t)',
          ADD + 't.SetSourceData ws.Range("A1:C3"), xlColumns: ws.Activate: Set u = wb.Charts.Add(Before:=t): r = Seen(u)',
          'ws.Range("A1:C3").Select: Set t = wb.Charts.Add(Before:=ws, Count:=3): r = Seen(wb.Charts(1)) & " / " & Seen(wb.Charts(3))',
          ADD + 'Set u = wb.Charts.Add(After:=wb.Sheets(wb.Sheets.Count)): r = Seen(u)', TWO + 'Set t = wb.Sheets.Add(Type:=xlChart, After:=ws): r = Seen(t)',
          ADD + 'Set u = wb.Charts.Add(After:=ws): r = Seen(u)', ADD + 'ws.Columns(1).Insert: ws.Activate: Set u = wb.Charts.Add(Before:=t): r = Seen(u)']

MAKE = '''    Set ws = wb.Worksheets(1)
    ws.Name = "One"
    wb.Worksheets.Add(After:=ws).Name = "Two"
    wb.Worksheets.Add(After:=wb.Worksheets("Two")).Name = "Three"
    ws.Activate
    ws.Range("A1:C1").Value = Array("", "Rev", "Cost")
    ws.Range("A2:C2").Value = Array("a", 10, 4)
    ws.Range("A3:C3").Value = Array("b", 20, 6)
    ws.Range("A1").Select
'''
DIGEST = ('wb.Sheets.Count', 'wb.Worksheets.Count', 'wb.Charts.Count', 'wb.ActiveSheet.Name', 'TypeName(wb.ActiveSheet)', 'TypeName(ActiveChart)')
SEEN = '''Private Function Seen(t As Object) As String
    Dim s As String, i As Long
    On Error Resume Next
    s = t.Name & " " & t.Index & " " & t.ChartType & " " & t.HasLegend & " " & t.HasTitle & " " & t.SeriesCollection.Count
    For i = 1 To t.SeriesCollection.Count
        s = s & " " & t.SeriesCollection(i).Formula
    Next
    If Err.Number Then s = s & " E" & Err.Number
    Seen = s
End Function
'''


def probe() -> str:
    lines = [SEEN, 'Private Function Read1(wb As Workbook) As String', '    Dim s As String, v As Variant, o As Object', '    On Error Resume Next']
    for expression in DIGEST:
        lines += ['    v = "E": Err.Clear', f'    v = {expression}', '    If Err.Number Then v = "E" & Err.Number', '    s = s & v & " "']
    lines += ['    Err.Clear', '    For Each o In wb.Sheets', '        s = s & "[" & o.Name & ":" & TypeName(o) & ":" & o.Index & "]"', '    Next', '    Err.Clear', '    Read1 = s', 'End Function', '',
              'Private Sub Act(ByVal which As Long, wb As Workbook, ws As Worksheet, r As Variant)', '    Dim t As Object, u As Object, co As Object', '    Select Case which']
    for index, statement in enumerate(CASES):
        lines += [f'    Case {index}', f'        {statement}']
    lines += ['    End Select', 'End Sub', '',
              'Private Function Tried(ByVal which As Long, ByVal folder As String) As String', '    Dim wb As Workbook, ws As Worksheet, r As Variant, s As String, v As String', '    Set wb = Workbooks.Add(xlWBATWorksheet)', MAKE.rstrip('\n'),
              '    On Error GoTo Failed', '    Act which, wb, ws, r', '    s = "ok "', 'Answered:', '    On Error Resume Next', '    Err.Clear', '    v = "": v = CStr(r)', '    If Err.Number Then v = "E" & Err.Number',
              '    s = s & "|" & v & "|" & Read1(wb)', '    Err.Clear', '    wb.Unprotect: ws.Unprotect', '    Err.Clear', '    wb.SaveAs folder & "\\c" & which & ".xlsx", xlOpenXMLWorkbook', '    If Err.Number Then s = s & "|saveE" & Err.Number',
              '    On Error GoTo 0', '    Tried = s', '    wb.Close False', '    Exit Function', 'Failed:', '    s = "Err " & Err.Number & " "', '    Resume Answered', 'End Function', '',
              'Public Function Probe(ByVal folder As String, ByVal first As Long, ByVal last As Long) As String', '    Dim i As Long, s As String', '    Application.DisplayAlerts = False', '    For i = first To last',
              '        s = s & "@@@" & Tried(i, folder)', '    Next', '    Application.DisplayAlerts = True', '    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def parts(path: Path) -> dict[str, str]:
    if not path.exists():
        return {}
    keep = re.compile(r'\[Content_Types\]\.xml|xl/workbook\.xml|xl/_rels/workbook\.xml\.rels|xl/chartsheets/.*|xl/drawings/.*|xl/charts/chart\d+\.xml|xl/charts/_rels/.*|docProps/app\.xml')
    with zipfile.ZipFile(path) as package:
        return {name: re.sub(r'<mc:AlternateContent[^>]*><mc:Choice Requires="x15"><x15ac:absPath.*?</mc:AlternateContent>', '',
                             package.read(name).decode('utf-8'), flags=re.S)
                for name in package.namelist() if keep.fullmatch(name)}


def batch(source: str, folder: str, first: int, last: int) -> list[str]:
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        done = excel.run_vba(source, 'Probe', (folder, first, last), timeout=280.0, line_numbers=False)
        if done.ok and not done.dialogs:
            return str(done.value).split('@@@')
    if first == last:
        return [f'STOPPED {done.outcome}']
    return [answer for index in range(first, last + 1) for answer in batch(source, folder, index, index)]


def main() -> None:
    source = probe()
    answers: list[str] = []
    with tempfile.TemporaryDirectory() as folder:
        for first in range(0, len(CASES), 25):
            answers += batch(source, folder, first, min(first + 24, len(CASES) - 1))
        saved = [parts(Path(folder) / f'c{index}.xlsx') for index in range(len(CASES))]
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source,
              'cases': [[statement, answer, part] for statement, answer, part in zip(CASES, answers, saved, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for statement, answer, _ in record['cases']:
        print(statement[:110], '=>', answer)


if __name__ == '__main__':
    main()
