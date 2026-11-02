"""An axis's Gridlines and TickLabels: what they read before and after each change, the errors, and the chart part
Excel saves.

Each case makes a workbook, adds a column chart through ChartObjects.Add, fills categories in A2:A4 and two series in
B and C, gives them to the chart with SetSourceData, and runs the case's statements with ax and vx as the category and
value axes. It then reads READS, each as its value (TypeName for an object) or "E" and the error number, and saves
the workbook. The chart part is recorded as scripts/measure_chart_markup.py records it.

    python scripts/measure_chart_axis_parts.py

writes tests/fixtures/chart_axis_parts.json, which tests/test_excel_chart_axis_parts.py replays.
"""
from __future__ import annotations

import json
import re
import tempfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

from measure_chart_markup import chart_part

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/chart_axis_parts.json'
CASES = [
    '',
    'vx.MajorGridlines.Delete',
    'Set gl = ax.MajorGridlines',
    'Set gl = vx.MinorGridlines',
    'ax.HasMajorGridlines = True: ax.MajorGridlines.Delete',
    'vx.HasMinorGridlines = True: vx.MinorGridlines.Delete',
    'ws.Range("B2:C4").NumberFormat = "0.00"',
    'vx.TickLabels.NumberFormat = "0%"',
    'vx.TickLabels.NumberFormat = "0%": vx.TickLabels.NumberFormatLinked = True',
    'vx.TickLabels.NumberFormatLinked = False',
    'ax.TickLabels.NumberFormat = "@"',
    'vx.TickLabels.NumberFormatLocal = "0.0"',
    'vx.TickLabels.NumberFormat = "General"',
    'vx.TickLabels.NumberFormat = "nonsense"',
    'ax.HasMajorGridlines = True: vx.HasMinorGridlines = True: ax.HasMinorGridlines = True',
    'vx.HasMinorGridlines = True: vx.HasMajorGridlines = False',
]
READS = [
    'vx.HasMajorGridlines', 'vx.HasMinorGridlines', 'ax.HasMajorGridlines', 'ax.HasMinorGridlines',
    'vx.TickLabels.NumberFormat', 'vx.TickLabels.NumberFormatLinked', 'vx.TickLabels.NumberFormatLocal',
    'ax.TickLabels.NumberFormat', 'ax.TickLabels.NumberFormatLinked',
    'vx.TickLabels.Name', 'ax.TickLabels.Name', 'vx.TickLabels.Orientation', 'ax.TickLabels.Orientation',
    'vx.TickLabels.Offset', 'ax.TickLabels.Offset', 'vx.TickLabels.Alignment', 'ax.TickLabels.Alignment',
    'vx.TickLabels.Depth', 'ax.TickLabels.Depth', 'vx.TickLabels.MultiLevel', 'ax.TickLabels.MultiLevel',
    'vx.TickLabels.ReadingOrder', 'TypeName(vx.TickLabels.Parent)',
    'TypeName(vx.TickLabels.Application)',
    'vx.MajorGridlines.Name', 'TypeName(vx.MajorGridlines.Parent)', 'TypeName(vx.MajorGridlines.Application)',
    'vx.HasMajorGridlines', 'vx.MinorGridlines.Name', 'ax.MajorGridlines.Name', 'ax.MinorGridlines.Name',
]


def _probe() -> str:
    lines = ['Private Function Tried(ByVal which As Long, ByVal folder As String) As String',
             '    Dim wb As Workbook, ws As Worksheet, ch As Object, ax As Object, vx As Object, gl As Object',
             '    Dim s As String',
             '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set ws = wb.Worksheets(1)',
             '    Set ch = ws.ChartObjects.Add(100, 10, 300, 200).Chart',
             '    ws.Range("A1:A4").Value = Application.Transpose(Array("k", "a", "b", "c"))',
             '    ws.Range("B1:B4").Value = Application.Transpose(Array("v", 1, 2, 3))',
             '    ws.Range("C1:C4").Value = Application.Transpose(Array("w", 4, 5, 6))',
             '    ch.SetSourceData ws.Range("A1:C4")', '    ch.ChartType = xlColumnClustered',
             '    Set ax = ch.Axes(xlCategory): Set vx = ch.Axes(xlValue)',
             '    On Error GoTo Failed', '    Select Case which']
    for index, statements in enumerate(CASES):
        lines += [f'    Case {index}'] + ([f'        {statements}'] if statements else [])
    lines += ['    End Select', '    On Error Resume Next']
    for read in READS:
        lines += ['    Err.Clear', f'    s = s & "," & {read}', '    If Err.Number Then s = s & "E" & Err.Number']
    lines += ['    On Error GoTo Failed', '    Tried = "ok " & Mid(s, 2)', 'Saved:',
              '    wb.SaveAs folder & "\\c" & which & ".xlsx", xlOpenXMLWorkbook', '    wb.Close False',
              '    Exit Function', 'Failed:', '    Tried = "Err " & Err.Number & " " & s', '    Resume Saved',
              'End Function', '',
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
              'cases': [[statements, answer, part]
                        for statements, answer, part in zip(CASES, answers, parts, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for case, answer, part in record['cases']:
        axes = re.findall(r'<c:(?:catAx|valAx)>.*?</c:(?:catAx|valAx)>', part)
        print(case, '->', answer)
        for axis in axes:
            print('    ', re.sub(r'<c:txPr>.*?</c:txPr>|<c:spPr>.*?</c:spPr>', '', axis))


if __name__ == '__main__':
    main()
