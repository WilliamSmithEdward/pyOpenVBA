"""Chart.Legend and its Position, IncludeInLayout, Shadow and Delete: what a legend reads before and after each change,
the errors bad values raise, and the chart part Excel saves.

Each case makes a workbook, adds a chart through ChartObjects.Add, fills categories in A2:A4 and two series in B and
C, gives them to the chart with SetSourceData and runs the case's statements, then saves it. The chart part is
recorded as scripts/measure_chart_markup.py records it.

    python scripts/measure_chart_legend.py

writes tests/fixtures/chart_legend.json, which tests/test_excel_chart_legend.py replays.
"""
from __future__ import annotations

import json
import tempfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

from measure_chart_markup import chart_part

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/chart_legend.json'
READ = ('"|" & ch.HasLegend & "," & ch.Legend.Position & "," & ch.Legend.IncludeInLayout & "," & ch.Legend.Shadow & '
        '"," & TypeName(ch.Legend.Parent)')
CASES = [
    '',
    'ch.Legend.Position = xlLegendPositionBottom',
    'ch.Legend.Position = xlLegendPositionTop',
    'ch.Legend.Position = xlLegendPositionLeft',
    'ch.Legend.Position = xlLegendPositionCorner',
    'ch.Legend.Position = xlLegendPositionRight',
    'ch.Legend.Position = xlLegendPositionCustom',
    'ch.Legend.Position = 5',
    'ch.Legend.IncludeInLayout = False',
    'ch.Legend.Shadow = True',
    'ch.Legend.Shadow = True: ch.Legend.Shadow = False',
    'ch.Legend.Delete',
    'ch.HasLegend = False: s = s & "|" & ch.HasLegend: s = s & "|" & TypeName(ch.Legend)',
    'ch.Legend.Position = xlLegendPositionTop: ch.HasLegend = False: ch.HasLegend = True',
    'ch.Legend.Delete: ch.HasLegend = True',
    'ch.Legend.Position = xlLegendPositionBottom: ch.Legend.IncludeInLayout = False: ch.Legend.Shadow = True',
    'ch.Legend.Position = xlLegendPositionLeft: ch.ChartType = xlPie',
]


def _probe() -> str:
    lines = ['Private Function Tried(ByVal which As Long, ByVal folder As String) As String',
             '    Dim wb As Workbook, ws As Worksheet, ch As Object, s As String',
             '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set ws = wb.Worksheets(1)',
             '    Set ch = ws.ChartObjects.Add(100, 10, 300, 200).Chart',
             '    ws.Range("A1:A4").Value = Application.Transpose(Array("k", "a", "b", "c"))',
             '    ws.Range("B1:B4").Value = Application.Transpose(Array("v", 1, 2, 3))',
             '    ws.Range("C1:C4").Value = Application.Transpose(Array("w", 4, 0.5, 6))',
             '    ch.SetSourceData ws.Range("A1:C4")', '    On Error GoTo Failed', f'    s = {READ}',
             '    Select Case which']
    for index, statements in enumerate(CASES):
        lines += [f'    Case {index}'] + ([f'        {statements}'] if statements else [])
    lines += ['    End Select', f'    If ch.HasLegend Then s = s & {READ}', '    Tried = "ok" & s', 'Saved:',
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
    for case, answer, part in record['cases']:
        import re
        print(case, '->', answer, '||', ' '.join(re.findall(r'<c:legend>.*?</c:legend>|<c:autoTitleDeleted[^>]*>', part)))


if __name__ == '__main__':
    main()
