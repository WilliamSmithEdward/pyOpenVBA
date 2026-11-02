"""Series.DataLabels: what the labels read on each kind of chart, what each setting changes, the errors, and the chart
part Excel saves.

Each case makes a workbook, adds a chart through ChartObjects.Add, fills categories in A2:A4 and two series in B and
C, gives them to the chart with SetSourceData, sets the case's type, turns the first series' labels on unless the case
says otherwise, and runs its statements with sr as that series and dl as its labels. It then reads READS, each as its
value (TypeName for an object) or "E" and the error number, and saves the workbook. The chart part is recorded as
scripts/measure_chart_markup.py records it.

    python scripts/measure_data_labels_object.py

writes tests/fixtures/data_labels_object.json, which tests/test_excel_data_labels_object.py replays.
"""
from __future__ import annotations

import json
import re
import tempfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

from measure_chart_markup import chart_part

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/data_labels_object.json'
#: (chart type, whether the labels are turned on first, statements)
CASES = [
    (51, True, ''), (57, True, ''), (4, True, ''), (65, True, ''), (5, True, ''), (-4120, True, ''),
    (-4169, True, ''), (1, True, ''), (52, True, ''),
    (51, False, ''),
    (51, False, 'Set dl = sr.DataLabels'),
    (51, True, 'dl.ShowCategoryName = True'),
    (51, True, 'dl.ShowSeriesName = True'),
    (51, True, 'dl.ShowValue = False'),
    (51, True, 'dl.ShowLegendKey = True'),
    (5, True, 'dl.ShowPercentage = True'),
    (51, True, 'dl.ShowPercentage = True'),
    (51, True, 'dl.ShowBubbleSize = True'),
    (51, True, 'dl.ShowValue = False: dl.ShowCategoryName = False'),
    (51, True, 'dl.ShowCategoryName = True: dl.ShowSeriesName = True: dl.ShowValue = False'),
    (51, True, 'dl.Position = xlLabelPositionInsideEnd'),
    (51, True, 'dl.Position = xlLabelPositionCenter'),
    (51, True, 'dl.Position = xlLabelPositionInsideBase'),
    (51, True, 'dl.Position = xlLabelPositionOutsideEnd'),
    (51, True, 'dl.Position = xlLabelPositionAbove'),
    (4, True, 'dl.Position = xlLabelPositionAbove'),
    (4, True, 'dl.Position = xlLabelPositionBelow'),
    (4, True, 'dl.Position = xlLabelPositionLeft'),
    (4, True, 'dl.Position = xlLabelPositionRight'),
    (5, True, 'dl.Position = xlLabelPositionBestFit'),
    (5, True, 'dl.Position = xlLabelPositionOutsideEnd'),
    (1, True, 'dl.Position = xlLabelPositionCenter'),
    (51, True, 'dl.NumberFormat = "0.0"'),
    (51, True, 'dl.NumberFormat = "0.0": dl.NumberFormatLinked = True'),
    (51, True, 'dl.NumberFormatLinked = False'),
    (51, True, 'dl.Separator = "; "'),
    (51, True, 'dl.ShowCategoryName = True: dl.Separator = "; "'),
    (51, True, 'dl.Delete'),
    (51, True, 'sr.HasDataLabels = False'),
    (51, True, 'dl.ShowValue = False: dl.ShowCategoryName = True: sr.HasDataLabels = False: sr.HasDataLabels = True'),
]
READS = [
    'sr.HasDataLabels', 'TypeName(sr.DataLabels)', 'dl.Count', 'dl.Name', 'TypeName(dl.Parent)', 'dl.ShowValue',
    'dl.ShowCategoryName', 'dl.ShowSeriesName', 'dl.ShowPercentage', 'dl.ShowLegendKey', 'dl.ShowBubbleSize',
    'dl.Position', 'dl.NumberFormat', 'dl.NumberFormatLinked', 'dl.Separator', 'dl.AutoText', 'dl.ShowRange',
    'dl.Orientation', 'dl.HorizontalAlignment', 'dl.VerticalAlignment', 'dl.ReadingOrder', 'dl.Shadow',
    'TypeName(dl(1))', 'TypeName(dl.Item(2))', 'dl(4).Name',
]


def _probe() -> str:
    lines = ['Private Function Tried(ByVal which As Long, ByVal folder As String) As String',
             '    Dim wb As Workbook, ws As Worksheet, ch As Object, sr As Object, dl As Object, s As String',
             '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set ws = wb.Worksheets(1)',
             '    Set ch = ws.ChartObjects.Add(100, 10, 300, 200).Chart',
             '    ws.Range("A1:A4").Value = Application.Transpose(Array("k", "a", "b", "c"))',
             '    ws.Range("B1:B4").Value = Application.Transpose(Array("v", 1, 2, 3))',
             '    ws.Range("C1:C4").Value = Application.Transpose(Array("w", 4, 5, 6))',
             '    ch.SetSourceData ws.Range("A1:C4")', '    On Error GoTo Failed', '    Select Case which']
    for index, (kind, on, statements) in enumerate(CASES):
        lines += [f'    Case {index}', f'        ch.ChartType = {kind}', '        Set sr = ch.SeriesCollection(1)']
        if on:
            lines += ['        sr.HasDataLabels = True', '        Set dl = sr.DataLabels']
        lines += [f'        {statements}'] if statements else []
    lines += ['    End Select', '    On Error Resume Next', '    Set dl = sr.DataLabels']
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
        result = excel.run_vba(source, 'Probe', (folder,), timeout=290.0)
        assert result.ok and not result.dialogs, result
        answers = str(result.value).split('@@@')
        parts = [chart_part(Path(folder) / f'c{index}.xlsx') for index in range(len(CASES))]
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'reads': READS, 'source': source,
              'cases': [[f'{kind} {"on" if on else "off"}: {statements}', answer, part]
                        for (kind, on, statements), answer, part in zip(CASES, answers, parts, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for case, answer, part in record['cases']:
        labels = re.search(r'<c:dLbls>.*?</c:dLbls>', part)
        print(case, '->', answer)
        print('    ', re.sub(r'<c:spPr>.*?</c:spPr>|<c:txPr>.*?</c:txPr>|<c:extLst>.*?</c:extLst>', '',
                             labels.group(0)) if labels else '')


if __name__ == '__main__':
    main()
