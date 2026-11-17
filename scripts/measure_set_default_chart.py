"""Measure Chart.SetDefaultChart: what it answers, and what Charts.Add makes after it, with a chart type and with xlBuiltIn,
always ending with SetDefaultChart xlBuiltIn. It sets Excel's default chart for every later session, so it is run only under
a registry snapshot that restores what the run added (William's approval, 2026-11-17); this script records the answers alone.
Needs Excel:

    python scripts/measure_set_default_chart.py [control]

writes tests/fixtures/set_default_chart.json; with "control" it runs the same session without SetDefaultChart and writes
nothing, to learn what any session changes.
"""

from __future__ import annotations

import datetime
import json
import sys
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

FIXTURES = Path(__file__).resolve().parents[1] / 'tests/fixtures'
SOURCE = r'''Public Function Probe(ByVal measure As Boolean) As String
    Dim wb As Workbook, ws As Worksheet, c As Object, made As Object, s As String
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Worksheets(1)
    ws.Range("A1:C1").Value = Array("", "Rev", "Cost")
    ws.Range("A2:C2").Value = Array("a", 10, 4)
    ws.Range("A3:C3").Value = Array("b", 20, 6)
    ws.Range("A1").Select
    Set c = ws.ChartObjects.Add(48, 96, 240, 144).Chart
    c.SetSourceData ws.Range("A1:C3")
    s = "before " & TypeName(c.ChartType) & " " & c.ChartType
    On Error Resume Next
    If measure Then
        Err.Clear
        c.SetDefaultChart xlLine
        s = s & "|set line " & Err.Number
        Err.Clear
        ws.Activate
        Set made = wb.Charts.Add
        s = s & "|add " & Err.Number & " " & made.ChartType
        Err.Clear
        c.SetDefaultChart xlBuiltIn
        s = s & "|set builtin " & Err.Number
        Err.Clear
        ws.Activate
        Set made = wb.Charts.Add
        s = s & "|add " & Err.Number & " " & made.ChartType
    End If
    ' Whatever happened above, the default goes back to the built-in one.
    Err.Clear
    If measure Then c.SetDefaultChart xlBuiltIn
    s = s & "|final " & Err.Number
    On Error GoTo 0
    Application.DisplayAlerts = False
    wb.Close False
    Application.DisplayAlerts = True
    Probe = s
End Function
'''


def main() -> None:
    measure = sys.argv[1:] != ['control']
    with ExcelSession(HarnessConfig(lock_wait_s=30)) as excel:
        done = excel.run_vba(SOURCE, 'Probe', (measure,), timeout=120.0, line_numbers=False)
    answer = done.value if done.outcome == 'passed' else f'{done.outcome}: {done.message}'
    print('answer:', answer, flush=True)
    if measure:
        record = {'measured_at': datetime.datetime.now(datetime.timezone.utc).isoformat(), 'source': SOURCE, 'answer': answer}
        (FIXTURES / 'set_default_chart.json').write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')


if __name__ == '__main__':
    main()
