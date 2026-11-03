"""One point's Format and ClearFormats on line and scatter charts, and a point's Delete and Select on every kind: what each
does, what it reads afterwards, and the chart part Excel saves.

Each case makes a workbook with a chart of categories in A2:A4 and two series in B and C, of the chart type TYPES names,
runs the case's statements, reads READS, each as its value or "E" and an error number, and saves the workbook. The chart
part is recorded as scripts/measure_chart_markup.py records it.

    python scripts/measure_point_line_format.py

writes tests/fixtures/point_line_format.json, which tests/test_excel_point_line_format.py replays.
"""
from __future__ import annotations

import json
import re
import tempfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

from measure_chart_markup import chart_part

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/point_line_format.json'
TYPES = {'markers': 'xlLineMarkers', 'line': 'xlLine', 'scatter': 'xlXYScatter', 'column': 'xlColumnClustered', 'area': 'xlArea', 'pie': 'xlPie'}
LINES = ('markers', 'line', 'scatter')
P, Q, S = 'ch.SeriesCollection(1).Points(2)', 'ch.SeriesCollection(1).Points(1)', 'ch.SeriesCollection(1)'
FILL, LINE = f'{P}.Format.Fill', f'{P}.Format.Line'
LOOKS = [
    f'{FILL}.ForeColor.RGB = RGB(255, 0, 0)', f'{FILL}.Visible = msoFalse', f'{FILL}.Visible = msoTrue', f'{FILL}.Solid', f'{FILL}.Transparency = 0.5',
    f'{LINE}.ForeColor.RGB = RGB(255, 0, 0)', f'{LINE}.Visible = msoFalse', f'{LINE}.Visible = msoTrue', f'{LINE}.Weight = 3', f'{LINE}.DashStyle = msoLineDash', f'{LINE}.Transparency = 0.5',
    f'{FILL}.ForeColor.RGB = RGB(255, 0, 0): {LINE}.ForeColor.RGB = RGB(0, 0, 255)',
    f'{LINE}.ForeColor.RGB = RGB(0, 0, 255): {LINE}.Weight = 3',
    f'{FILL}.ForeColor.RGB = RGB(255, 0, 0): {FILL}.Visible = msoFalse',
    f'{LINE}.ForeColor.RGB = RGB(255, 0, 0): {LINE}.Visible = msoFalse',
    f'{P}.MarkerBackgroundColor = RGB(0, 255, 0): {FILL}.ForeColor.RGB = RGB(255, 0, 0)',
    f'{P}.MarkerForegroundColor = RGB(0, 255, 0): {LINE}.ForeColor.RGB = RGB(255, 0, 0)',
    f'{FILL}.ForeColor.RGB = RGB(255, 0, 0): {P}.MarkerBackgroundColor = RGB(0, 255, 0)',
    f'{LINE}.ForeColor.RGB = RGB(255, 0, 0): {P}.MarkerForegroundColor = RGB(0, 255, 0)',
    f'{S}.Format.Fill.ForeColor.RGB = RGB(0, 255, 0): {FILL}.ForeColor.RGB = RGB(255, 0, 0)',
    f'{S}.Format.Line.ForeColor.RGB = RGB(0, 255, 0): {LINE}.ForeColor.RGB = RGB(255, 0, 0)',
    f'{FILL}.ForeColor.RGB = RGB(255, 0, 0): {S}.Format.Fill.ForeColor.RGB = RGB(0, 255, 0)',
    f'{LINE}.ForeColor.RGB = RGB(255, 0, 0): {S}.Format.Line.ForeColor.RGB = RGB(0, 255, 0)',
    f'{P}.ClearFormats', f'{FILL}.ForeColor.RGB = RGB(255, 0, 0): {P}.ClearFormats', f'{LINE}.ForeColor.RGB = RGB(255, 0, 0): {P}.ClearFormats',
    f'{P}.MarkerSize = 12: {P}.ClearFormats', f'{P}.MarkerBackgroundColor = RGB(0, 255, 0): {P}.ClearFormats', f'{P}.ClearFormats: {P}.ClearFormats',
    f'{S}.Format.Line.ForeColor.RGB = RGB(0, 255, 0): {P}.ClearFormats',
]
ANY_KIND = [
    f'{P}.Delete', f'{Q}.Delete', f'ch.SeriesCollection(2).Points(2).Delete', f'{P}.Delete: {P}.Delete',
    f'{P}.Select', f'{Q}.Select', f'ch.Parent.Activate: {P}.Select', f'ch.Parent.Activate: ch.SeriesCollection(2).Points(3).Select',
]
CASES = [(kind, statement) for kind in LINES for statement in LOOKS]
CASES += [(kind, f'{P}.ClearFormats') for kind in ('area', 'pie')]
CASES += [(kind, statement) for kind in TYPES for statement in ANY_KIND]
READS = [f'{FILL}.ForeColor.RGB', f'{FILL}.Visible', f'{FILL}.Transparency', f'{LINE}.ForeColor.RGB', f'{LINE}.Visible', f'{LINE}.Weight', f'{LINE}.DashStyle', f'{LINE}.Transparency',
         f'{P}.MarkerStyle', f'{P}.MarkerSize', f'{P}.MarkerBackgroundColor', f'{P}.MarkerForegroundColor', f'{Q}.MarkerBackgroundColor', f'{Q}.MarkerForegroundColor',
         f'{S}.MarkerBackgroundColor', f'{S}.MarkerForegroundColor', f'{S}.Format.Line.ForeColor.RGB', f'{Q}.Format.Line.ForeColor.RGB',
         'ch.SeriesCollection.Count', f'{S}.Points.Count', 'TypeName(Selection)', 'Selection.Name']


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
