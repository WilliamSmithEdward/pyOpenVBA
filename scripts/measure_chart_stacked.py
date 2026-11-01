"""Stacked and 100% stacked column, bar, line and area charts: what the chart and its first series read, the errors,
and the chart part Excel saves.

Each case makes a workbook, adds a chart through ChartObjects.Add, fills categories in A2:A4 and two series in B and
C, gives them to the chart with SetSourceData, sets the case's chart type and runs the case's statements, then reads
the chart's ChartType, the first series' ChartType and Type, and the chart group's Overlap and GapWidth, each as its
value or "E" and the error number, and saves it. The chart part is recorded as scripts/measure_chart_markup.py
records it.

    python scripts/measure_chart_stacked.py

writes tests/fixtures/chart_stacked.json, which tests/test_excel_chart_stacked.py replays.
"""
from __future__ import annotations

import json
import re
import tempfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

from measure_chart_markup import chart_part

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/chart_stacked.json'
CASES = [
    ('xlColumnStacked', ''),
    ('xlColumnStacked100', ''),
    ('xlBarStacked', ''),
    ('xlBarStacked100', ''),
    ('xlLineStacked', ''),
    ('xlLineStacked100', ''),
    ('xlLineMarkersStacked', ''),
    ('xlLineMarkersStacked100', ''),
    ('xlAreaStacked', ''),
    ('xlAreaStacked100', ''),
    ('xlColumnStacked', 'ch.ChartType = xlColumnClustered'),
    ('xlColumnClustered', 'ch.ChartType = xlColumnStacked: ch.ChartType = xlColumnClustered'),
    ('xlColumnStacked', 'ch.ChartGroups(1).Overlap = 50'),
    ('xlColumnStacked', 'ch.ChartGroups(1).GapWidth = 80'),
    ('xlLineStacked', 'ch.SeriesCollection(1).MarkerStyle = xlMarkerStyleCircle'),
    ('xlLineMarkersStacked', 'ch.SeriesCollection(1).MarkerStyle = xlMarkerStyleNone'),
    ('xlColumnStacked', 'ch.SeriesCollection.NewSeries'),
]
READS = ['ch.ChartType', 'ch.SeriesCollection(1).ChartType', 'ch.SeriesCollection(1).Type',
         'ch.ChartGroups(1).Overlap', 'ch.ChartGroups(1).GapWidth']


def _probe() -> str:
    lines = ['Private Function Tried(ByVal which As Long, ByVal folder As String) As String',
             '    Dim wb As Workbook, ws As Worksheet, ch As Object, s As String, i As Long',
             '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set ws = wb.Worksheets(1)',
             '    Set ch = ws.ChartObjects.Add(100, 10, 300, 200).Chart',
             '    ws.Range("A1:A4").Value = Application.Transpose(Array("k", "a", "b", "c"))',
             '    ws.Range("B1:B4").Value = Application.Transpose(Array("v", 1, 2, 3))',
             '    ws.Range("C1:C4").Value = Application.Transpose(Array("w", 4, 5, 6))',
             '    ch.SetSourceData ws.Range("A1:C4")', '    On Error GoTo Failed', '    Select Case which']
    for index, (kind, statements) in enumerate(CASES):
        lines += [f'    Case {index}', f'        ch.ChartType = {kind}']
        lines += [f'        {statements}'] if statements else []
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
              'cases': [[f'{kind}: {statements}', answer, part]
                        for (kind, statements), answer, part in zip(CASES, answers, parts, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for case, answer, part in record['cases']:
        group = re.search(r'<c:plotArea>.*?<c:ser>', part)
        print(case, '->', answer, '||', re.sub(r'<c:layout/>', '', group.group(0)) if group else '')


if __name__ == '__main__':
    main()
