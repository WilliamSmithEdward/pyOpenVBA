"""Font on a chart's text: what the title, axis title, tick labels, legend, data labels and chart area read, what each
change does, the errors, and the chart part Excel saves.

Each case makes a workbook, adds a column chart through ChartObjects.Add of categories in A2:A4 and two series in B and
C, with a title, a legend, a value axis title and the first series' labels on, gives the element named el the case's
text if it says so, runs the change on el.Font, reads READS, each as its value (TypeName for an object) or "E" and the
error number, and saves the workbook. The chart part is recorded as scripts/measure_chart_markup.py records it.

    python scripts/measure_chart_fonts.py

writes tests/fixtures/chart_fonts.json, which tests/test_excel_chart_fonts.py replays.
"""
from __future__ import annotations

import json
import re
import tempfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

from measure_chart_markup import chart_part

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/chart_fonts.json'
ELEMENTS = {
    'title': 'ch.ChartTitle', 'typed': 'ch.ChartTitle', 'axistitle': 'ch.Axes(xlValue).AxisTitle',
    'axistyped': 'ch.Axes(xlValue).AxisTitle', 'ticks': 'ch.Axes(xlValue).TickLabels',
    'catticks': 'ch.Axes(xlCategory).TickLabels', 'legend': 'ch.Legend',
    'labels': 'ch.SeriesCollection(1).DataLabels', 'area': 'ch.ChartArea',
}
CHANGES = ['', 'el.Font.Size = 14', 'el.Font.Bold = True', 'el.Font.Italic = True', 'el.Font.Color = 255',
           'el.Font.Name = "Arial"', 'el.Font.Underline = xlUnderlineStyleSingle', 'el.Font.Strikethrough = True',
           'el.Font.ColorIndex = 3', 'el.Font.Size = 14: el.Font.Bold = True: el.Font.Color = 255',
           'el.Font.Bold = True: el.Font.Bold = False', 'el.Font.Size = 0', 'el.Font.Size = 500',
           'el.Font.FontStyle = "Bold Italic"', 'el.Font.Superscript = True', 'el.Font.Name = ""']
CASES = [(element, change) for element in ELEMENTS for change in CHANGES]
READS = ['TypeName(el.Font)', 'TypeName(el.Font.Parent)', 'TypeName(el.Font.Application)', 'el.Font.Name',
         'el.Font.Size', 'el.Font.Bold', 'el.Font.Italic', 'el.Font.Underline', 'el.Font.Strikethrough',
         'el.Font.Color', 'el.Font.ColorIndex', 'el.Font.FontStyle', 'el.Font.Superscript', 'el.Font.Subscript',
         'el.Font.Shadow', 'el.Font.OutlineFont', 'el.Font.Background', 'el.Font.ThemeColor', 'el.Font.TintAndShade',
         'el.Font.ThemeFont']


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
             '    ch.SeriesCollection(1).HasDataLabels = True', '    On Error GoTo Failed', '    Select Case which']
    for index, (element, change) in enumerate(CASES):
        lines += [f'    Case {index}', f'        Set el = {ELEMENTS[element]}']
        lines += ['        el.Text = "T"'] if element in ('typed', 'axistyped') else []
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
              '    Application.DisplayAlerts = False', '    For i = first To last',
              '        s = s & "@@@" & Tried(i, folder)', '    Next', '    Application.DisplayAlerts = True',
              '    Probe = Mid(s, 4)', 'End Function']
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
    for case, answer, part in record['cases']:
        found = re.findall(r'<c:(?:rich|txPr)>.*?</c:(?:rich|txPr)>', part, re.DOTALL)
        print(case, '\n  =', answer, '\n  +', ' || '.join(found))


if __name__ == '__main__':
    main()
