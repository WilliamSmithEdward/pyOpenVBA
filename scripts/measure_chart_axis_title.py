"""AxisTitle.Formula and its variants, Orientation, HorizontalAlignment, VerticalAlignment, ReadingOrder, Position,
IncludeInLayout and Shadow: what an axis title reads before and after each change, the errors bad values raise, and
the chart part Excel saves.

Each case makes a workbook, adds a chart through ChartObjects.Add, fills categories in A2:A4 and two series in B and
C, gives them to the chart with SetSourceData, gives the axis the case names a title and runs the case's statements
on it, then saves it. The chart part is recorded as scripts/measure_chart_markup.py records it.

    python scripts/measure_chart_axis_title.py

writes tests/fixtures/chart_axis_title.json, which tests/test_excel_chart_axis_title.py replays.
"""
from __future__ import annotations

import json
import tempfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

from measure_chart_markup import chart_part

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/chart_axis_title.json'
READ = ('"|" & a.HasTitle & "," & t.Text & "," & t.Caption & "," & t.Formula & "," & t.FormulaR1C1 & "," & '
        't.FormulaLocal & "," & t.FormulaR1C1Local & "," & t.Orientation & "," & t.HorizontalAlignment & "," & '
        't.VerticalAlignment & "," & t.ReadingOrder & "," & t.Position & "," & t.IncludeInLayout & "," & t.Shadow')
#: (axis, statements)
CASES = [
    (1, ''),
    (2, ''),
    (1, 't.Caption = "Days"'),
    (2, 't.Formula = "=Sheet1!$C$1"'),
    (1, 't.Formula = "=Sheet1!A1"'),
    (2, 't.FormulaR1C1 = "=Sheet1!R1C2"'),
    (1, 't.Formula = "=1+2"'),
    (2, 't.Formula = "plain"'),
    (1, 't.Orientation = xlUpward'),
    (2, 't.Orientation = xlHorizontal'),
    (2, 't.Orientation = xlDownward'),
    (2, 't.Orientation = 45'),
    (1, 't.Orientation = 91'),
    (1, 't.HorizontalAlignment = xlLeft'),
    (2, 't.HorizontalAlignment = xlLeft'),
    (2, 't.Orientation = xlHorizontal: t.HorizontalAlignment = xlRight'),
    (1, 't.HorizontalAlignment = 99'),
    (1, 't.VerticalAlignment = xlTop'),
    (2, 't.ReadingOrder = xlRTL'),
    (1, 't.IncludeInLayout = False'),
    (2, 't.Shadow = True'),
    (1, 't.Caption = "Up": t.Orientation = xlUpward'),
    (2, 't.Caption = "Units": t.ReadingOrder = xlLTR'),
    (2, 't.Delete'),
    (1, 't.Position = xlChartElementPositionAutomatic'),
]


def _probe() -> str:
    lines = ['Private Function Tried(ByVal which As Long, ByVal folder As String) As String',
             '    Dim wb As Workbook, ws As Worksheet, ch As Object, a As Object, t As Object, s As String',
             '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set ws = wb.Worksheets(1)',
             '    Set ch = ws.ChartObjects.Add(100, 10, 300, 200).Chart',
             '    ws.Range("A1:A4").Value = Application.Transpose(Array("k", "a", "b", "c"))',
             '    ws.Range("B1:B4").Value = Application.Transpose(Array("v", 1, 2, 3))',
             '    ws.Range("C1:C4").Value = Application.Transpose(Array("w", 4, 0.5, 6))',
             '    ch.SetSourceData ws.Range("A1:C4")', '    On Error GoTo Failed', '    Select Case which']
    for index, (axis, statements) in enumerate(CASES):
        lines += [f'    Case {index}', f'        Set a = ch.Axes({axis}): a.HasTitle = True: Set t = a.AxisTitle',
                  f'        s = {READ}'] + ([f'        {statements}'] if statements else [])
    lines += ['    End Select', '    If a.HasTitle Then Set t = a.AxisTitle: s = s & ' + READ,
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
              'cases': [[f'{axis}: {statements}', answer, part]
                        for (axis, statements), answer, part in zip(CASES, answers, parts, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for case, answer, _ in record['cases']:
        print(case, '->', answer.split('|', 2)[0], answer.split('|')[-1])


if __name__ == '__main__':
    main()
