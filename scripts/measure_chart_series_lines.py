"""Series.FormulaR1C1, FormulaLocal, FormulaR1C1Local, Smooth, MarkerStyle and MarkerSize: what a series reads before
and after each change, the errors bad values raise, and the chart part Excel saves.

Each case makes a workbook, adds a chart through ChartObjects.Add, fills categories in A2:A4 and two series in B and
C, gives them to the chart with SetSourceData, sets its type and runs the case's statements on its first series, then
saves it. The chart part is recorded as scripts/measure_chart_markup.py records it.

    python scripts/measure_chart_series_lines.py

writes tests/fixtures/chart_series_lines.json, which tests/test_excel_chart_series_lines.py replays.
"""
from __future__ import annotations

import json
import tempfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

from measure_chart_markup import chart_part

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/chart_series_lines.json'
READ = ('"|" & ch.ChartType & "," & x.Smooth & "," & x.MarkerStyle & "," & x.MarkerSize & "," & '
        'ch.SeriesCollection(2).MarkerStyle & "," & ch.SeriesCollection(2).MarkerSize')
#: Nine series of literal values, the first two replacing those SetSourceData made.
NINE = ('ch.SeriesCollection(2).Delete: ch.SeriesCollection(1).Delete: '
        'For i = 1 To 9: ch.SeriesCollection.NewSeries.Values = Array(i, i + 1): Next: '
        'For i = 1 To 9: ch.SeriesCollection(i).MarkerSize = 7: s = s & ";" & ch.SeriesCollection(i).MarkerStyle: Next')
FORMULAS = '"|" & x.Formula & "|" & x.FormulaR1C1 & "|" & x.FormulaLocal & "|" & x.FormulaR1C1Local'
#: (chart type, statements)
CASES = [
    ('xlColumnClustered', 's = s & ' + FORMULAS),
    ('xlLine', 's = s & ' + FORMULAS),
    ('xlLineMarkers', ''),
    ('xlXYScatter', ''),
    ('xlPie', ''),
    ('xlLine', 'x.Smooth = True'),
    ('xlLineMarkers', 'x.Smooth = True'),
    ('xlXYScatter', 'x.Smooth = True'),
    ('xlLine', 'x.Smooth = True: x.Smooth = False'),
    ('xlLine', 'x.Smooth = "x"'),
    ('xlLine', 'x.MarkerStyle = xlMarkerStyleCircle'),
    ('xlLineMarkers', 'x.MarkerStyle = xlMarkerStyleSquare'),
    ('xlLineMarkers', 'x.MarkerStyle = xlMarkerStyleNone'),
    ('xlLineMarkers', 'x.MarkerStyle = xlMarkerStyleAutomatic'),
    ('xlXYScatter', 'x.MarkerStyle = xlMarkerStyleTriangle'),
    ('xlLineMarkers', 'x.MarkerStyle = 99'),
    ('xlLineMarkers', 'x.MarkerSize = 9'),
    ('xlLine', 'x.MarkerSize = 9'),
    ('xlXYScatter', 'x.MarkerSize = 2: x.MarkerStyle = xlMarkerStyleDiamond'),
    ('xlLineMarkers', 'x.MarkerSize = 1'),
    ('xlLineMarkers', 'x.MarkerSize = 73'),
    ('xlLineMarkers', 'x.MarkerSize = 7.6'),
    ('xlLineMarkers', 'x.MarkerSize = 5'),
    ('xlLineMarkers', 'x.MarkerStyle = xlMarkerStyleX: x.MarkerSize = 12: ch.ChartType = xlLine'),
    ('xlColumnClustered', 'x.MarkerStyle = xlMarkerStyleCircle'),
    ('xlColumnClustered', 'x.Smooth = True'),
    ('xlXYScatter', 'x.MarkerStyle = xlMarkerStylePicture'),
    ('xlLineMarkers', 'x.MarkerStyle = xlMarkerStyleStar'),
    ('xlLineMarkers', 'x.MarkerStyle = 0'),
    ('xlLineMarkers', 'x.MarkerStyle = 6'),
    ('xlLineMarkers', 'x.MarkerStyle = 4'),
    ('xlLineMarkers', 'x.MarkerStyle = 7'),
    ('xlLineMarkers', 'x.MarkerStyle = 10'),
    ('xlLineMarkers', 'x.MarkerStyle = -1'),
    ('xlLineMarkers', 'x.MarkerStyle = xlMarkerStylePlus'),
    ('xlLineMarkers', 'x.MarkerStyle = xlMarkerStyleDot'),
    ('xlLineMarkers', 'x.MarkerStyle = xlMarkerStyleDash'),
    ('xlLineMarkers', 'x.MarkerStyle = xlMarkerStyleX'),
    ('xlLine', 'x.MarkerStyle = xlMarkerStyleAutomatic'),
    ('xlLine', 'ch.SeriesCollection(2).MarkerStyle = xlMarkerStyleCircle'),
    ('xlLineMarkers', 'ch.SeriesCollection(2).MarkerStyle = xlMarkerStyleNone'),
    ('xlLine', 'x.MarkerSize = 11: x.MarkerStyle = xlMarkerStyleSquare'),
    ('xlLineMarkers', 'x.Smooth = True: ch.ChartType = xlLine'),
    ('xlLine', 'x.MarkerStyle = xlMarkerStyleCircle: ch.ChartType = xlLineMarkers'),
    ('xlLineMarkers', 'x.MarkerStyle = xlMarkerStyleCircle: ch.ChartType = xlXYScatter'),
    ('xlLineMarkers', 'x.MarkerSize = 1.5'),
    ('xlLineMarkers', 'x.MarkerSize = 72'),
    ('xlLineMarkers', NINE),
]


def _probe() -> str:
    lines = ['Private Function Tried(ByVal which As Long, ByVal folder As String) As String',
             '    Dim wb As Workbook, ws As Worksheet, ch As Object, x As Object, s As String, i As Long',
             '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set ws = wb.Worksheets(1)',
             '    Set ch = ws.ChartObjects.Add(100, 10, 300, 200).Chart',
             '    ws.Range("A1:A4").Value = Application.Transpose(Array("k", "a", "b", "c"))',
             '    ws.Range("B1:B4").Value = Application.Transpose(Array("v", 1, 2, 3))',
             '    ws.Range("C1:C4").Value = Application.Transpose(Array("w", 4, 0.5, 6))',
             '    ch.SetSourceData ws.Range("A1:C4")', '    On Error GoTo Failed', '    Select Case which']
    for index, (kind, statements) in enumerate(CASES):
        lines += [f'    Case {index}', f'        ch.ChartType = {kind}: Set x = ch.SeriesCollection(1)',
                  f'        s = {READ}']
        if statements:
            lines.append(f'        {statements}')
        lines.append(f'        s = s & {READ}')
    lines += ['    End Select', '    Tried = "ok" & s', 'Saved:',
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
              'cases': [[f'{kind}: {statements}', answer, part]
                        for (kind, statements), answer, part in zip(CASES, answers, parts, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for case, answer, _ in record['cases']:
        print(case, '->', answer)


if __name__ == '__main__':
    main()
