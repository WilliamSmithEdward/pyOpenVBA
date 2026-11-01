"""ChartTitle.Caption, Delete, Formula and its variants, Orientation, HorizontalAlignment, VerticalAlignment,
ReadingOrder, Position, IncludeInLayout and Shadow: what a title reads before and after each change, the errors bad
values raise, and the chart part Excel saves.

Each case makes a workbook, adds a chart through ChartObjects.Add, fills categories in A2:A4 and two series in B and
C, gives them to the chart with SetSourceData, gives it a title and runs the case's statements on it, then saves it.
The chart part is recorded as scripts/measure_chart_markup.py records it.

    python scripts/measure_chart_title.py

writes tests/fixtures/chart_title.json, which tests/test_excel_chart_title.py replays.
"""
from __future__ import annotations

import json
import tempfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

from measure_chart_markup import chart_part

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/chart_title.json'
READ = ('"|" & ch.HasTitle & "," & t.Text & "," & t.Caption & "," & t.Formula & "," & t.FormulaR1C1 & "," & '
        't.FormulaLocal & "," & t.FormulaR1C1Local & "," & t.Orientation & "," & t.HorizontalAlignment & "," & '
        't.VerticalAlignment & "," & t.ReadingOrder & "," & t.Position & "," & t.IncludeInLayout & "," & t.Shadow')
CASES = [
    '',
    't.Caption = "Hello"',
    't.Text = "Typed": t.Caption = "Again"',
    't.Formula = "=Sheet1!$B$1"',
    't.Formula = "=Sheet1!B1"',
    't.FormulaR1C1 = "=Sheet1!R1C3"',
    't.Formula = "=Sheet1!$B$1": ws.Range("B1").Value = "changed"',
    't.Formula = "=Sheet1!$B$1": t.Text = "plain"',
    't.Formula = "plain text"',
    't.Formula = "=1+2"',
    't.Formula = ""',
    't.Orientation = xlUpward',
    't.Orientation = xlDownward',
    't.Orientation = xlVertical',
    't.Orientation = 45',
    't.Orientation = -90',
    't.Orientation = 91',
    't.HorizontalAlignment = xlLeft',
    't.HorizontalAlignment = xlRight',
    't.HorizontalAlignment = xlJustify',
    't.HorizontalAlignment = xlDistributed',
    't.HorizontalAlignment = 99',
    't.VerticalAlignment = xlTop',
    't.VerticalAlignment = xlBottom',
    't.ReadingOrder = xlRTL',
    't.IncludeInLayout = False',
    't.Shadow = True',
    't.Delete',
    't.Delete: s = s & "|" & ch.HasTitle',
    't.Caption = "x": ch.HasTitle = False: ch.HasTitle = True',
    't.Caption = "Up": t.Orientation = xlUpward',
    't.Orientation = xlUpward: t.Caption = "Up"',
    't.Formula = "=Sheet1!$B$1": t.HorizontalAlignment = xlLeft',
    't.Orientation = 30: t.HorizontalAlignment = xlRight: t.ReadingOrder = xlRTL',
    't.Shadow = True: t.IncludeInLayout = False: t.Caption = "Both"',
    't.Delete: ch.HasTitle = True',
    't.Orientation = 90: t.Orientation = xlHorizontal: t.HorizontalAlignment = xlCenter: t.ReadingOrder = xlContext',
    't.Shadow = True: t.Shadow = False',
    't.ReadingOrder = xlLTR',
    't.Position = xlChartElementPositionAutomatic',
    't.Orientation = 30: t.HorizontalAlignment = xlRight',
    't.HorizontalAlignment = xlRight: t.ReadingOrder = xlRTL',
    't.HorizontalAlignment = xlRight: t.Orientation = 30',
    't.Caption = "Left": t.HorizontalAlignment = xlLeft: t.ReadingOrder = xlLTR',
    't.Orientation = xlUpward: t.HorizontalAlignment = xlLeft',
]


def _probe() -> str:
    lines = ['Private Function Tried(ByVal which As Long, ByVal folder As String) As String',
             '    Dim wb As Workbook, ws As Worksheet, ch As Object, t As Object, s As String',
             '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set ws = wb.Worksheets(1)',
             '    Set ch = ws.ChartObjects.Add(100, 10, 300, 200).Chart',
             '    ws.Range("A1:A4").Value = Application.Transpose(Array("k", "a", "b", "c"))',
             '    ws.Range("B1:B4").Value = Application.Transpose(Array("v", 1, 2, 3))',
             '    ws.Range("C1:C4").Value = Application.Transpose(Array("w", 4, 0.5, 6))',
             '    ch.SetSourceData ws.Range("A1:C4")', '    ch.HasTitle = True', '    Set t = ch.ChartTitle',
             '    On Error GoTo Failed', f'    s = {READ}', '    Select Case which']
    for index, statements in enumerate(CASES):
        lines += [f'    Case {index}'] + ([f'        {statements}'] if statements else [])
    lines += ['    End Select', '    If ch.HasTitle Then Set t = ch.ChartTitle: s = s & ' + READ,
              '    Tried = "ok" & s', 'Saved:',
              '    wb.SaveAs folder & "\\c" & which & ".xlsx", xlOpenXMLWorkbook', '    wb.Close False',
              '    Exit Function', 'Failed:', '    Tried = "Err " & Err.Number & s', '    Resume Saved', 'End Function', '',
              'Public Function Probe(ByVal folder As String) As String', '    Dim i As Long, s As String',
              '    Application.DisplayAlerts = False', f'    For i = 0 To {len(CASES) - 1}',
              '        s = s & "@@@" & Tried(i, folder)', '    Next', '    Application.DisplayAlerts = True',
              '    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def main() -> None:
    source = _probe()
    with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(source, 'Probe', (folder,), timeout=280.0)
        assert result.ok and not result.dialogs, result
        answers = str(result.value).split('@@@')
        parts = [chart_part(Path(folder) / f'c{index}.xlsx') for index in range(len(CASES))]
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source,
              'cases': [[statements, answer, part] for statements, answer, part in zip(CASES, answers, parts, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for case, answer, _ in record['cases']:
        print(case, '->', answer)


if __name__ == '__main__':
    main()
