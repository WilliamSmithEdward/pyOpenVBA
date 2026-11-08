"""Charts.Add placed before a chart sheet read from a file, which keeps no source range: what the new chart plots.

Each case makes a workbook with One holding a small table in A1:C3 and a chart sheet Mine plotting the case's range,
saves it, closes it, opens it again, runs the case's statements, and records what the new chart plots and the tabs.
The rivals, written before measuring: (L1) the block the loaded chart's series cover, in the orientation it plots them,
(L2) nothing, (L3) what the worksheet has selected.

    python scripts/measure_charts_add_loaded.py

writes tests/fixtures/charts_add_loaded.json, which tests/test_excel_charts_add.py replays.
"""
from __future__ import annotations

import json
import tempfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig
from pyvbaharness.results import HarnessError

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/charts_add_loaded.json'
#: Each case: the range the chart sheet plots and how, the cell selected on One before the save, and what runs after the reopen.
CASES: list[tuple[str, str, str]] = [
    ('ws.Range("A1:B3")', 'A1', 'Set t = wb.Charts.Add(Before:=wb.Charts("Mine")): r = Seen(t)'),
    ('ws.Range("A1:B3")', 'E10', 'Set t = wb.Charts.Add(Before:=wb.Charts("Mine")): r = Seen(t)'),
    ('ws.Range("A1:B3")', 'A1', 'wb.Charts("Mine").Activate: Set t = wb.Charts.Add: r = Seen(t)'),
    ('ws.Range("A1:C3"), xlColumns', 'E10', 'Set t = wb.Charts.Add(Before:=wb.Charts("Mine")): r = Seen(t)'),
    ('ws.Range("A1:C3"), xlRows', 'E10', 'Set t = wb.Charts.Add(Before:=wb.Charts("Mine")): r = Seen(t)'),
    ('ws.Range("A1:A3,C1:C3")', 'E10', 'Set t = wb.Charts.Add(Before:=wb.Charts("Mine")): r = Seen(t)'),
    ('ws.Range("B1:C3")', 'E10', 'Set t = wb.Charts.Add(Before:=wb.Charts("Mine"), Count:=2): r = Seen(t) & " / " & Seen(wb.Charts(2))'),
    ('ws.Range("A1:B3")', 'E10', 'Set t = wb.Charts.Add(After:=wb.Worksheets("One")): r = Seen(t)'),
]
# Only a loaded chart sheet activated in the session plotted again, what its series cover. The same shapes activated, and the
# rivals for that: (C1) the block its series cover, in its orientation, (C2) the block guessed afresh, (C3) its series' own areas.
ACTIVATED = 'wb.Charts("Mine").Activate: ws.Activate: Set t = wb.Charts.Add(Before:=wb.Charts("Mine")): r = Seen(t)'
CASES += [('ws.Range("A1:C3"), xlColumns', 'E10', ACTIVATED), ('ws.Range("A1:C3"), xlRows', 'E10', ACTIVATED), ('ws.Range("A1:A3,C1:C3")', 'E10', ACTIVATED),
          ('ws.Range("B1:C3")', 'E10', ACTIVATED), ('ws.Range("A1:B3")', 'A1', ACTIVATED),
          ('ws.Range("A1:B3")', 'E10', 'wb.Charts("Mine").Activate: Set t = wb.Charts.Add(Count:=2): r = Seen(t) & " / " & Seen(wb.Charts(2))'),
          ('ws.Range("A1:C3"), xlRows', 'E10', 'wb.Charts("Mine").Activate: wb.Charts("Mine").SeriesCollection(1).Delete: Set t = wb.Charts.Add: r = Seen(t)')]
# Held out: the rule from the cases above, the cells its series take now with their label corner, as columns, on cases written after it.
CASES += [('ws.Range("A1:C3"), xlColumns', 'E10', 'wb.Charts("Mine").Activate: ws.Activate: ws.Rows(1).Insert: Set t = wb.Charts.Add(Before:=wb.Charts("Mine")): r = Seen(t)'),
          ('ws.Range("B1:C3")', 'E10', 'wb.Charts("Mine").Activate: Set t = wb.Charts.Add: r = Seen(t)'),
          ('ws.Range("A1:B3")', 'E10', 'wb.Charts("Mine").Activate: ws.Activate: Set t = wb.Charts.Add(After:=ws): r = Seen(t)'),
          ('ws.Range("A1:C3"), xlRows', 'E10', 'wb.Charts("Mine").Activate: Set t = wb.Charts.Add(Before:=wb.Charts("Mine")): Set u = wb.Charts.Add(Before:=wb.Charts("Mine")): r = Seen(u)')]
# Held out again: what its series take is fixed at the address they had when it was first activated, an insert after that moving
# nothing (case 15). An insert before the first activation, and one between a first and a second.
CASES += [('ws.Range("A1:B3")', 'E10', 'ws.Rows(1).Insert: wb.Charts("Mine").Activate: ws.Activate: Set t = wb.Charts.Add(Before:=wb.Charts("Mine")): r = Seen(t)'),
          ('ws.Range("A1:B3")', 'E10', 'wb.Charts("Mine").Activate: ws.Activate: ws.Columns(1).Insert: wb.Charts("Mine").Activate: ws.Activate: '
                                       'Set t = wb.Charts.Add(Before:=wb.Charts("Mine")): r = Seen(t)')]
# Held out a third time: what its series take is read again at each activation (case 20), and only then (case 15).
CASES += [('ws.Range("A1:C3"), xlRows', 'E10', 'wb.Charts("Mine").Activate: ws.Activate: ws.Rows(1).Insert: wb.Charts("Mine").Activate: '
                                                 'Set t = wb.Charts.Add: r = Seen(t)'),
          ('ws.Range("A1:B3"), xlColumns', 'E10', 'wb.Charts("Mine").Activate: ws.Activate: ws.Columns(1).Insert: '
                                                    'Set t = wb.Charts.Add(Before:=wb.Charts("Mine")): r = Seen(t)')]
SOURCE = '''Private Function Seen(t As Object) As String
    Dim s As String, i As Long
    On Error Resume Next
    s = t.Name & " " & t.Index & " " & t.ChartType & " " & t.HasLegend & " " & t.HasTitle & " " & t.SeriesCollection.Count
    For i = 1 To t.SeriesCollection.Count
        s = s & " " & t.SeriesCollection(i).Formula
    Next
    If Err.Number Then s = s & " E" & Err.Number
    Seen = s
End Function

Private Function Tabs(wb As Workbook) As String
    Dim o As Object, s As String
    For Each o In wb.Sheets
        s = s & "[" & o.Name & ":" & TypeName(o) & "]"
    Next
    Tabs = s & " " & wb.ActiveSheet.Name
End Function

Public Function Probe(ByVal path As String, ByVal plotted As String, ByVal selected As String, ByVal which As Long) As String
    Dim wb As Workbook, ws As Worksheet, t As Object, u As Object, r As Variant, c As Object
    Application.DisplayAlerts = False
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Worksheets(1)
    ws.Name = "One"
    ws.Range("A1:C1").Value = Array("", "Rev", "Cost")
    ws.Range("A2:C2").Value = Array("a", 10, 4)
    ws.Range("A3:C3").Value = Array("b", 20, 6)
    ws.Range("E10").Select
    Set c = wb.Charts.Add
    c.Name = "Mine"
    Select Case plotted
@PLOTTED@
    End Select
    ws.Activate
    ws.Range(selected).Select
    wb.SaveAs path, xlOpenXMLWorkbook
    wb.Close False
    Set wb = Workbooks.Open(path)
    Set ws = wb.Worksheets("One")
    On Error GoTo Failed
    Select Case which
@CASES@
    End Select
    Probe = "ok |" & r & "|" & Tabs(wb)
    GoTo Done
Failed:
    Probe = "Err " & Err.Number & "|" & Tabs(wb)
    Resume Done
Done:
    wb.Close False
    Application.DisplayAlerts = True
End Function
'''


def source() -> str:
    plotted = sorted({case[0] for case in CASES})
    return (SOURCE.replace('@PLOTTED@', '\n'.join(f'    Case "{index}"\n        c.SetSourceData {range_}' for index, range_ in enumerate(plotted)))
            .replace('@CASES@', '\n'.join(f'    Case {index}\n        {case[2]}' for index, case in enumerate(CASES))))


def plotted_key(range_: str) -> str:
    return str(sorted({case[0] for case in CASES}).index(range_))


def main() -> None:
    code = source()
    answers: list[str] = []
    with tempfile.TemporaryDirectory() as folder:
        for index, (range_, selected, _) in enumerate(CASES):
            # A session of its own for each case, so one that stops Excel stops only itself.
            try:
                with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
                    excel.new_document()
                    done = excel.run_vba(code, 'Probe', (str(Path(folder) / f'l{index}.xlsx'), plotted_key(range_), selected, index),
                                         timeout=120.0, line_numbers=False)
                    answers.append(str(done.value) if done.ok and not done.dialogs else f'STOPPED {done.outcome}')
            except HarnessError as failure:
                answers.append(f'STOPPED {failure}')
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': code,
              'cases': [[*case, answer] for case, answer in zip(CASES, answers, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for case, answer in zip(CASES, answers, strict=True):
        print(case, '=>', answer)


if __name__ == '__main__':
    main()
