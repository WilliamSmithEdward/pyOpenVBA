"""Series.HasErrorBars, Series.ErrorBar and the ErrorBars object: what they read before and after each change, the
errors, and the chart part Excel saves.

Each case makes a workbook, adds a chart through ChartObjects.Add, fills categories in A2:A4 and two series in B and
C, gives them to the chart with SetSourceData, sets the case's chart type and runs the case's statements with sr as
the first series, then reads HasErrorBars and, where there are error bars, ErrorBars' EndStyle and Name, each as its
value or "E" and the error number, and saves it. The chart part is recorded as scripts/measure_chart_markup.py
records it.

    python scripts/measure_chart_error_bars.py

writes tests/fixtures/chart_error_bars.json, which tests/test_excel_chart_error_bars.py replays.
"""
from __future__ import annotations

import json
import re
import tempfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

from measure_chart_markup import chart_part

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/chart_error_bars.json'
CASES = [
    ('xlColumnClustered', ''),
    ('xlColumnClustered', 'sr.HasErrorBars = True'),
    ('xlColumnClustered', 'sr.HasErrorBars = True: sr.HasErrorBars = False'),
    ('xlColumnClustered', 'sr.ErrorBar xlY, xlErrorBarIncludeBoth, xlErrorBarTypeFixedValue, 1'),
    ('xlColumnClustered', 'sr.ErrorBar xlY, xlErrorBarIncludePlusValues, xlErrorBarTypeFixedValue, 0.5'),
    ('xlColumnClustered', 'sr.ErrorBar xlY, xlErrorBarIncludeMinusValues, xlErrorBarTypePercent, 10'),
    ('xlColumnClustered', 'sr.ErrorBar xlY, xlErrorBarIncludeBoth, xlErrorBarTypeStDev, 2'),
    ('xlColumnClustered', 'sr.ErrorBar xlY, xlErrorBarIncludeBoth, xlErrorBarTypeStError'),
    ('xlColumnClustered', 'sr.ErrorBar xlY, xlErrorBarIncludeBoth, xlErrorBarTypeCustom, 1, 2'),
    ('xlColumnClustered', 'sr.ErrorBar xlY, xlErrorBarIncludeNone, xlErrorBarTypeFixedValue, 1'),
    ('xlColumnClustered', 'sr.ErrorBar xlX, xlErrorBarIncludeBoth, xlErrorBarTypeFixedValue, 1'),
    ('xlColumnClustered', 'sr.ErrorBar xlY, xlErrorBarIncludeBoth, xlErrorBarTypeFixedValue, 1: '
                          'sr.ErrorBars.EndStyle = xlNoCap'),
    ('xlColumnClustered', 'sr.ErrorBar xlY, xlErrorBarIncludeBoth, xlErrorBarTypeFixedValue, 1: sr.ErrorBars.Delete'),
    ('xlColumnClustered', 'sr.ErrorBar xlY, xlErrorBarIncludeBoth, xlErrorBarTypeFixedValue, -1'),
    ('xlColumnClustered', 'sr.ErrorBar xlY, 7, xlErrorBarTypeFixedValue, 1'),
    ('xlColumnClustered', 'sr.ErrorBar xlY, xlErrorBarIncludeBoth, 9, 1'),
    ('xlBarClustered', 'sr.HasErrorBars = True'),
    ('xlLine', 'sr.HasErrorBars = True'),
    ('xlXYScatter', 'sr.HasErrorBars = True'),
    ('xlXYScatter', 'sr.ErrorBar xlX, xlErrorBarIncludeBoth, xlErrorBarTypeFixedValue, 1'),
    ('xlPie', 'sr.HasErrorBars = True'),
    ('xlColumnStacked', 'sr.HasErrorBars = True'),
    # HasErrorBars takes its amount from the value axis' major unit (scripts/measure_chart_error_bars.py's notes).
    ('xlColumnClustered', 'ch.Axes(xlValue).MajorUnit = 2: sr.HasErrorBars = True'),
    ('xlBarClustered', 'ch.Axes(xlValue).MajorUnit = 0.25: sr.HasErrorBars = True'),
    ('xlLine', 'ch.Axes(xlValue).MajorUnit = 3: sr.HasErrorBars = True'),
    ('xlXYScatter', 'ch.Axes(xlValue).MajorUnit = 3: sr.HasErrorBars = True'),
    ('xlPie', 'ws.Range("B2:B4").Value = Application.Transpose(Array(100, 250, 700)): sr.HasErrorBars = True'),
    ('xlDoughnut', 'sr.HasErrorBars = True'),
    ('xlColumnClustered', 'ch.SeriesCollection(2).ErrorBar xlY, xlErrorBarIncludeBoth, xlErrorBarTypeFixedValue, 1: '
                          's = ch.SeriesCollection(2).ErrorBars.Name & ";"'),
    ('xlColumnClustered', 'sr.ErrorBar xlY, xlErrorBarIncludeBoth, xlErrorBarTypeFixedValue, 1: '
                          'sr.ErrorBar xlY, xlErrorBarIncludePlusValues, xlErrorBarTypePercent, 5'),
    ('xlColumnClustered', 'sr.ErrorBar xlY, xlErrorBarIncludeBoth, xlErrorBarTypeFixedValue, 1: '
                          'sr.ErrorBars.EndStyle = 7'),
    ('xlColumnClustered', 'sr.ErrorBar xlY, xlErrorBarIncludeBoth, xlErrorBarTypeFixedValue, 1: '
                          's = TypeName(sr.ErrorBars.Parent) & ";"'),
    ('xlColumnClustered', 'sr.ErrorBar xlY, xlErrorBarIncludeBoth, xlErrorBarTypeCustom, '
                          'ws.Range("C2:C4"), ws.Range("B2:B4")'),
    ('xlColumnClustered', 'sr.ErrorBar xlY, xlErrorBarIncludeBoth, xlErrorBarTypeCustom, Array(1, 2, 3), 1'),
    ('xlColumnClustered', 'sr.ErrorBar xlY, xlErrorBarIncludeBoth, xlErrorBarTypeFixedValue'),
    ('xlXYScatter', 'sr.ErrorBar xlY, xlErrorBarIncludeBoth, xlErrorBarTypeFixedValue, 1: '
                    'sr.ErrorBar xlX, xlErrorBarIncludePlusValues, xlErrorBarTypeFixedValue, 2'),
    ('xlXYScatter', 'sr.ErrorBar xlX, xlErrorBarIncludeBoth, xlErrorBarTypeFixedValue, 1: sr.HasErrorBars = False'),
]
READS = ['sr.HasErrorBars', 'sr.ErrorBars.EndStyle', 'sr.ErrorBars.Name']


def _probe() -> str:
    lines = ['Private Function Tried(ByVal which As Long, ByVal folder As String) As String',
             '    Dim wb As Workbook, ws As Worksheet, ch As Object, sr As Object, s As String',
             '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set ws = wb.Worksheets(1)',
             '    Set ch = ws.ChartObjects.Add(100, 10, 300, 200).Chart',
             '    ws.Range("A1:A4").Value = Application.Transpose(Array("k", "a", "b", "c"))',
             '    ws.Range("B1:B4").Value = Application.Transpose(Array("v", 1, 2, 3))',
             '    ws.Range("C1:C4").Value = Application.Transpose(Array("w", 4, 5, 6))',
             '    ch.SetSourceData ws.Range("A1:C4")', '    Set sr = ch.SeriesCollection(1)',
             '    On Error GoTo Failed', '    Select Case which']
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
        bars = re.findall(r'<c:errBars>.*?</c:errBars>', part)
        print(case, '->', answer, '||', [re.sub(r'<c:spPr>.*?</c:spPr>', '<spPr/>', bar) for bar in bars])


if __name__ == '__main__':
    main()
