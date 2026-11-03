"""One point's marker: what setting each member does on each kind of chart, what it reads afterwards, and the chart part
Excel saves.

Each case makes a workbook with a chart of categories in A2:A4 and two series in B and C, of the chart type TYPES names,
runs the case's statements, reads READS, each as its value or "E" and an error number, and saves the workbook. The chart
part is recorded as scripts/measure_chart_markup.py records it.

    python scripts/measure_point_markers.py

writes tests/fixtures/point_markers.json, which tests/test_excel_point_markers.py replays.
"""
from __future__ import annotations

import json
import re
import tempfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

from measure_chart_markup import chart_part

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/point_markers.json'
TYPES = {'markers': 'xlLineMarkers', 'line': 'xlLine', 'scatter': 'xlXYScatter', 'column': 'xlColumnClustered', 'area': 'xlArea', 'pie': 'xlPie'}
P, S = 'ch.SeriesCollection(1).Points(2)', 'ch.SeriesCollection(1)'
ONE = [
    f'{P}.MarkerStyle = {style}' for style in ('xlMarkerStyleCircle', 'xlMarkerStyleSquare', 'xlMarkerStyleDiamond', 'xlMarkerStyleTriangle', 'xlMarkerStyleX', 'xlMarkerStyleStar',
                                               'xlMarkerStyleDot', 'xlMarkerStyleDash', 'xlMarkerStylePlus', 'xlMarkerStyleNone', 'xlMarkerStyleAutomatic', 'xlMarkerStylePicture', '99')
] + [f'{P}.MarkerSize = {size}' for size in (2, 7, 72, 1, 73, 12.5)] + [
    f'{P}.{member} = {colour}' for member in ('MarkerBackgroundColor', 'MarkerForegroundColor') for colour in ('RGB(255, 0, 0)', '0', '-1', '-2', '16777216')
] + [
    f'{P}.{member} = {index}' for member in ('MarkerBackgroundColorIndex', 'MarkerForegroundColorIndex') for index in ('3', '1', '57', 'xlColorIndexNone', 'xlColorIndexAutomatic', '99')
]
TOGETHER = [
    f'{P}.MarkerStyle = xlMarkerStyleSquare: {P}.MarkerSize = 12: {P}.MarkerBackgroundColor = RGB(255, 0, 0): {P}.MarkerForegroundColor = RGB(0, 0, 255)',
    f'{P}.MarkerSize = 12: {P}.MarkerStyle = xlMarkerStyleSquare',
    f'{P}.MarkerBackgroundColor = RGB(255, 0, 0): {P}.MarkerStyle = xlMarkerStyleSquare',
    f'{P}.MarkerStyle = xlMarkerStyleSquare: {P}.MarkerStyle = xlMarkerStyleAutomatic',
    f'{P}.MarkerStyle = xlMarkerStyleSquare: {P}.MarkerStyle = xlMarkerStyleNone',
    f'{P}.MarkerBackgroundColor = RGB(255, 0, 0): {P}.MarkerBackgroundColor = -1',
    f'{P}.MarkerSize = 12: ch.SeriesCollection(1).Points(1).MarkerSize = 14',
    f'{S}.MarkerStyle = xlMarkerStyleDiamond: {P}.MarkerSize = 12',
    f'{S}.MarkerSize = 11: {P}.MarkerStyle = xlMarkerStyleSquare',
    f'{S}.MarkerBackgroundColor = RGB(0, 255, 0): {P}.MarkerStyle = xlMarkerStyleSquare',
    f'{S}.MarkerBackgroundColor = RGB(0, 255, 0): {P}.MarkerBackgroundColor = RGB(255, 0, 0)',
    f'{P}.MarkerStyle = xlMarkerStyleSquare: {S}.MarkerStyle = xlMarkerStyleDiamond',
    f'{P}.MarkerSize = 12: {S}.MarkerSize = 11',
    f'{P}.MarkerBackgroundColor = RGB(255, 0, 0): {S}.MarkerBackgroundColor = RGB(0, 255, 0)',
    f'{P}.MarkerStyle = xlMarkerStyleSquare: {P}.ClearFormats',
    f'{P}.MarkerStyle = xlMarkerStyleSquare: {P}.Format.Fill.ForeColor.RGB = RGB(255, 0, 0)',
    f'{P}.Format.Fill.ForeColor.RGB = RGB(255, 0, 0): {P}.MarkerStyle = xlMarkerStyleSquare',
    f'{P}.MarkerStyle = xlMarkerStyleSquare: ch.ChartType = xlColumnClustered',
]
CASES = [(kind, statement) for kind in TYPES for statement in ONE]
CASES += [(kind, statement) for kind in ('markers', 'line', 'scatter') for statement in TOGETHER]
ONE_POINT = ['MarkerStyle', 'MarkerSize', 'MarkerBackgroundColor', 'MarkerForegroundColor', 'MarkerBackgroundColorIndex', 'MarkerForegroundColorIndex']
READS = [f'{P}.{name}' for name in ONE_POINT] + [f'ch.SeriesCollection(1).Points(1).{name}' for name in ONE_POINT] + [f'{S}.{name}' for name in ONE_POINT]


def _probe() -> str:
    lines = ['Private Function Tried(ByVal which As Long, ByVal folder As String) As String',
             '    Dim wb As Workbook, ws As Worksheet, ch As Object, s As String, v As String', '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set ws = wb.Worksheets(1)',
             '    Set ch = ws.ChartObjects.Add(100, 10, 300, 200).Chart', '    ws.Range("A1:A4").Value = Application.Transpose(Array("k", "a", "b", "c"))',
             '    ws.Range("B1:B4").Value = Application.Transpose(Array("v", 1, 2, 3))', '    ws.Range("C1:C4").Value = Application.Transpose(Array("w", 4, 5, 6))',
             '    ch.SetSourceData ws.Range("A1:C4")', '    On Error GoTo Failed', '    Select Case which']
    for index, (kind, statement) in enumerate(CASES):
        lines += [f'    Case {index}', f'        ch.ChartType = {TYPES[kind]}', f'        {statement}']
    lines += ['    End Select', '    s = "ok "', 'Answered:', '    On Error Resume Next']
    for read in READS:
        lines += ['    Err.Clear', '    v = ""', f'    v = {read}', '    If Err.Number Then v = "E" & Err.Number', '    s = s & "," & v']
    lines += ['    On Error GoTo 0', '    Tried = s', '    wb.SaveAs folder & "\\c" & which & ".xlsx", xlOpenXMLWorkbook', '    wb.Close False', '    Exit Function',
              'Failed:', '    s = "Err " & Err.Number & " "', '    Resume Answered', 'End Function', '',
              'Public Function Probe(ByVal folder As String, ByVal first As Long, ByVal last As Long) As String', '    Dim i As Long, s As String',
              '    Application.DisplayAlerts = False', '    For i = first To last', '        s = s & "@@@" & Tried(i, folder)', '    Next', '    Application.DisplayAlerts = True',
              '    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def main() -> None:
    source = _probe()
    answers: list[str] = []
    with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        for first in range(0, len(CASES), 30):
            last = min(first + 29, len(CASES) - 1)
            result = excel.run_vba(source, 'Probe', (folder, first, last), timeout=280.0)
            assert result.ok and not result.dialogs, result
            answers += str(result.value).split('@@@')
        parts = [chart_part(Path(folder) / f'c{index}.xlsx') for index in range(len(CASES))]
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'reads': READS, 'source': source,
              'cases': [[f'{kind}: {statement}', answer, part] for (kind, statement), answer, part in zip(CASES, answers, parts, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for case, answer, part in record['cases']:
        found = re.findall(r'<c:dPt>.*?</c:dPt>', part, re.DOTALL)
        print(case, '\n  =', answer, '\n  +', ' || '.join(found)[:300])


if __name__ == '__main__':
    main()
