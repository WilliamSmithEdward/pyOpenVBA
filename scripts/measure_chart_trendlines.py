"""Series.Trendlines, Trendlines.Add and the Trendline object: what they read after each change, the errors, and the
chart part Excel saves.

Each case makes a workbook, adds a chart through ChartObjects.Add, fills categories in A2:A5 and two series in B and
C, gives them to the chart with SetSourceData, sets the case's chart type and runs the case's statements with sr as
the first series, then reads the series' trendline count and, for each trendline, its Type, Order, Period, Forward,
Backward, Name, DisplayEquation, DisplayRSquared and Index, each as its value or "E" and the error number, and saves
it. The chart part is recorded as scripts/measure_chart_markup.py records it.

    python scripts/measure_chart_trendlines.py

writes tests/fixtures/chart_trendlines.json, which tests/test_excel_chart_trendlines.py replays.
"""
from __future__ import annotations

import json
import re
import tempfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

from measure_chart_markup import chart_part

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/chart_trendlines.json'
CASES = [
    ('xlColumnClustered', ''),
    ('xlColumnClustered', 'sr.Trendlines.Add'),
    ('xlColumnClustered', 'sr.Trendlines.Add xlLinear'),
    ('xlColumnClustered', 'sr.Trendlines.Add xlExponential'),
    ('xlColumnClustered', 'sr.Trendlines.Add xlLogarithmic'),
    ('xlColumnClustered', 'sr.Trendlines.Add xlPower'),
    ('xlColumnClustered', 'sr.Trendlines.Add xlPolynomial'),
    ('xlColumnClustered', 'sr.Trendlines.Add xlPolynomial, 3'),
    ('xlColumnClustered', 'sr.Trendlines.Add xlMovingAvg'),
    ('xlColumnClustered', 'sr.Trendlines.Add xlMovingAvg, , 3'),
    ('xlColumnClustered', 'sr.Trendlines.Add xlLinear, , , 1, 0.5'),
    ('xlColumnClustered', 'sr.Trendlines.Add xlLinear, Name:="Fit"'),
    ('xlColumnClustered', 'sr.Trendlines.Add xlLinear, Intercept:=2'),
    ('xlColumnClustered', 'sr.Trendlines.Add xlLinear: sr.Trendlines.Add xlPolynomial, 2'),
    ('xlColumnClustered', 'sr.Trendlines.Add xlLinear: sr.Trendlines(1).Delete'),
    ('xlColumnClustered', 'sr.Trendlines.Add xlLinear: sr.Trendlines(1).Order = 4'),
    ('xlColumnClustered', 'sr.Trendlines.Add xlPolynomial, 2: sr.Trendlines(1).Order = 7'),
    ('xlColumnClustered', 'sr.Trendlines.Add xlPolynomial, 7'),
    ('xlColumnClustered', 'sr.Trendlines.Add xlMovingAvg, , 1'),
    ('xlColumnClustered', 'sr.Trendlines.Add xlMovingAvg, , 5'),
    ('xlColumnClustered', 'sr.Trendlines.Add 9'),
    ('xlColumnClustered', 'sr.Trendlines.Add xlLinear: sr.Trendlines(1).Type = xlLogarithmic'),
    ('xlColumnClustered', 'sr.Trendlines.Add xlLinear: sr.Trendlines(1).Name = "Mine"'),
    ('xlColumnClustered', 'sr.Trendlines.Add xlLinear: sr.Trendlines(1).Forward = 2'),
    ('xlColumnClustered', 'sr.Trendlines.Add xlLinear: s = TypeName(sr.Trendlines(1).Parent) & ";"'),
    ('xlColumnClustered', 'ch.SeriesCollection(2).Trendlines.Add xlLinear: s = ch.SeriesCollection(2).Trendlines(1).Name & ";"'),
    ('xlColumnClustered', 's = sr.Trendlines.Count & ";": sr.Trendlines(1).Delete'),
    ('xlLine', 'sr.Trendlines.Add xlLinear'),
    ('xlBarClustered', 'sr.Trendlines.Add xlLinear'),
    ('xlXYScatter', 'sr.Trendlines.Add xlLinear'),
    ('xlArea', 'sr.Trendlines.Add xlLinear'),
    ('xlPie', 'sr.Trendlines.Add xlLinear'),
    ('xlColumnStacked', 'sr.Trendlines.Add xlLinear'),
    ('xlColumnClustered', 'sr.Trendlines.Add xlMovingAvg, , 4'),
    ('xlColumnClustered', 'sr.Trendlines.Add xlPolynomial, 6'),
    ('xlXYScatter', 'sr.Trendlines.Add xlLinear, , , 1, 0.5'),
    ('xlColumnClustered', 'sr.Trendlines.Add xlLinear, Intercept:=2: s = sr.Trendlines(1).Intercept & ";" & '
                          'sr.Trendlines(1).InterceptIsAuto & ";"'),
    ('xlColumnClustered', 'sr.Trendlines.Add xlLinear: s = sr.Trendlines(1).InterceptIsAuto & ";" & '
                          'sr.Trendlines(1).DisplayEquation & ";" & sr.Trendlines(1).DisplayRSquared & ";"'),
    ('xlColumnClustered', 'sr.Trendlines.Add xlLinear, DisplayEquation:=True'),
    ('xlColumnClustered', 'sr.Trendlines.Add xlLinear: sr.Trendlines(1).DisplayRSquared = True'),
    ('xlColumnClustered', 'sr.Trendlines.Add xlLinear, Name:="Fit": sr.Trendlines(1).NameIsAuto = True'),
    ('xlColumnClustered', 'sr.Trendlines.Add xlLinear: sr.Trendlines(1).Period = 3'),
    ('xlColumnClustered', 'sr.Trendlines.Add xlMovingAvg: sr.Trendlines(1).Period = 3'),
    ('xlColumnClustered', 'sr.Trendlines.Add xlLinear: sr.Trendlines(1).Type = xlMovingAvg'),
    ('xlColumnClustered', 'sr.Trendlines.Add xlLinear: sr.Trendlines(1).Forward = -1'),
    ('xlColumnClustered', 'sr.Trendlines.Add xlLinear: sr.Trendlines(1).Backward = 1'),
]
READS = ['Type', 'Order', 'Period', 'Forward', 'Backward', 'Name', 'Index', 'NameIsAuto']


def _probe() -> str:
    lines = ['Private Function Tried(ByVal which As Long, ByVal folder As String) As String',
             '    Dim wb As Workbook, ws As Worksheet, ch As Object, sr As Object, s As String, i As Long',
             '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set ws = wb.Worksheets(1)',
             '    Set ch = ws.ChartObjects.Add(100, 10, 300, 200).Chart',
             '    ws.Range("A1:A5").Value = Application.Transpose(Array("k", "a", "b", "c", "d"))',
             '    ws.Range("B1:B5").Value = Application.Transpose(Array("v", 1, 3, 2, 5))',
             '    ws.Range("C1:C5").Value = Application.Transpose(Array("w", 4, 5, 7, 6))',
             '    ch.SetSourceData ws.Range("A1:C5")', '    Set sr = ch.SeriesCollection(1)',
             '    On Error GoTo Failed', '    Select Case which']
    for index, (kind, statements) in enumerate(CASES):
        lines += [f'    Case {index}', f'        ch.ChartType = {kind}']
        lines += [f'        {statements}'] if statements else []
    lines += ['    End Select', '    On Error Resume Next', '    s = s & sr.Trendlines.Count',
              '    For i = 1 To sr.Trendlines.Count', '        s = s & "|"']
    for read in READS:
        lines += ['        Err.Clear', f'        s = s & "," & sr.Trendlines(i).{read}',
                  '        If Err.Number Then s = s & "E" & Err.Number']
    lines += ['    Next', '    On Error GoTo Failed', '    Tried = "ok " & s', 'Saved:',
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
        lines = re.findall(r'<c:trendline>.*?</c:trendline>', part)
        print(case, '->', answer, '||', [re.sub(r'<c:spPr>.*?</c:spPr>', '<spPr/>', line) for line in lines])


if __name__ == '__main__':
    main()
