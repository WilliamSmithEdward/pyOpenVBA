"""Format on a chart's other elements: what the chart area, plot area, legend, title, axes, axis title and gridlines
read through Format.Fill and Format.Line before and after each change, and the chart part Excel saves.

Each case makes a workbook, adds a column chart through ChartObjects.Add of categories in A2:A4 and two series in B and
C, with a title, a legend and a value axis title, runs the change on the element named el, reads READS, each as its
value (TypeName for an object) or "E" and the error number, and saves the workbook. The chart part is recorded as
scripts/measure_chart_markup.py records it.

    python scripts/measure_chart_element_format.py

writes tests/fixtures/chart_element_format.json, which tests/test_excel_chart_element_format.py replays.
"""
from __future__ import annotations

import json
import re
import tempfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

from measure_chart_markup import chart_part

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/chart_element_format.json'
ELEMENTS = {'area': 'ch.ChartArea', 'plot': 'ch.PlotArea', 'legend': 'ch.Legend', 'title': 'ch.ChartTitle',
            'value': 'ch.Axes(xlValue)', 'category': 'ch.Axes(xlCategory)', 'grid': 'ch.Axes(xlValue).MajorGridlines',
            'axistitle': 'ch.Axes(xlValue).AxisTitle'}
CHANGES = ['', 'el.Format.Fill.ForeColor.RGB = 255', 'el.Format.Fill.Visible = msoFalse',
           'el.Format.Fill.Transparency = 0.5', 'el.Format.Fill.Solid', 'el.Format.Line.ForeColor.RGB = 255',
           'el.Format.Line.Visible = msoFalse', 'el.Format.Line.Weight = 3', 'el.Format.Line.DashStyle = msoLineDash',
           'el.Format.Fill.ForeColor.RGB = 255: el.Format.Line.ForeColor.RGB = 65280: el.Format.Line.Weight = 2',
           'el.Format.Fill.ForeColor.RGB = 255: el.Format.Fill.Transparency = 0.5']
CASES = [(element, change) for element in ELEMENTS for change in CHANGES]
READS = ['TypeName(el.Format)', 'TypeName(el.Format.Parent)', 'el.Format.Fill.Visible', 'el.Format.Fill.ForeColor.RGB',
         'el.Format.Fill.Transparency', 'el.Format.Fill.Type', 'el.Format.Fill.BackColor.RGB',
         'el.Format.Fill.ForeColor.ObjectThemeColor', 'el.Format.Fill.ForeColor.Type', 'el.Format.Line.Visible',
         'el.Format.Line.ForeColor.RGB', 'el.Format.Line.Weight', 'el.Format.Line.DashStyle', 'el.Format.Line.Style',
         'el.Format.Line.Transparency', 'el.Format.Line.ForeColor.ObjectThemeColor', 'el.Format.Line.ForeColor.Type']


def _probe() -> str:
    lines = ['Private Function Tried(ByVal which As Long, ByVal folder As String) As String',
             '    Dim wb As Workbook, ws As Worksheet, ch As Object, el As Object, s As String, v As String',
             '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set ws = wb.Worksheets(1)',
             '    Set ch = ws.ChartObjects.Add(100, 10, 300, 200).Chart',
             '    ws.Range("A1:A4").Value = Application.Transpose(Array("k", "a", "b", "c"))',
             '    ws.Range("B1:B4").Value = Application.Transpose(Array("v", 1, 2, 3))',
             '    ws.Range("C1:C4").Value = Application.Transpose(Array("w", 4, 5, 6))',
             '    ch.SetSourceData ws.Range("A1:C4")', '    ch.ChartType = xlColumnClustered',
             '    ch.HasTitle = True: ch.HasLegend = True', '    ch.Axes(xlValue).HasTitle = True',
             '    On Error GoTo Failed', '    Select Case which']
    for index, (element, change) in enumerate(CASES):
        lines += [f'    Case {index}', f'        Set el = {ELEMENTS[element]}']
        lines += [f'        {change}'] if change else []
    lines += ['    End Select', '    s = "ok "', 'Answered:', '    On Error Resume Next']
    for read in READS:
        lines += ['    Err.Clear', '    v = ""', f'    v = {read}', '    If Err.Number Then v = "E" & Err.Number',
                  '    s = s & "," & v']
    lines += ['    On Error GoTo 0', '    Tried = s',
              '    wb.SaveAs folder & "\\c" & which & ".xlsx", xlOpenXMLWorkbook', '    wb.Close False',
              '    Exit Function', 'Failed:', '    s = "Err " & Err.Number & " "', '    Resume Answered',
              'End Function', '', 'Public Function Probe(ByVal folder As String, ByVal first As Long, '
              'ByVal last As Long) As String', '    Dim i As Long, s As String',
              '    Application.DisplayAlerts = False', '    For i = first To last', '        s = s & "@@@" & Tried(i, folder)', '    Next',
              '    Application.DisplayAlerts = True', '    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def main() -> None:
    source = _probe()
    answers: list[str] = []
    with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        for first in range(0, len(CASES), 40):
            last = min(first + 39, len(CASES) - 1)
            result = excel.run_vba(source, 'Probe', (folder, first, last), timeout=280.0)
            assert result.ok and not result.dialogs, result
            answers += str(result.value).split('@@@')
        parts = [chart_part(Path(folder) / f'c{index}.xlsx') for index in range(len(CASES))]
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'reads': READS, 'source': source,
              'cases': [[f'{element}: {change}', answer, part]
                        for (element, change), answer, part in zip(CASES, answers, parts, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    base = record['cases'][0][2]
    for case, answer, part in record['cases']:
        # Only what differs from the first case's part.
        start = next((i for i, (one, two) in enumerate(zip(base, part, strict=False)) if one != two), len(part))
        end = next((i for i, (one, two) in enumerate(zip(base[::-1], part[::-1], strict=False)) if one != two), 0)
        print(case, '\n  =', answer, '\n  +', part[max(0, start - 40):len(part) - end + 20] if part != base else '')
        if re.search(r'a14:', part):
            print('   (a14)')


if __name__ == '__main__':
    main()
