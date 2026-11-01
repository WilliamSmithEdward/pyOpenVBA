"""Series.HasDataLabels and HasLeaderLines: what they read before and
after each change, the errors, and the chart part Excel saves.

Each case makes a workbook, adds a chart through ChartObjects.Add, fills categories in A2:A4 and two series in B and
C, gives them to the chart with SetSourceData, sets the case's chart type and runs the case's statements with sr as
the first series, then saves it. The chart part is recorded as scripts/measure_chart_markup.py records it.

    python scripts/measure_chart_data_labels.py

writes tests/fixtures/chart_data_labels.json, which tests/test_excel_chart_data_labels.py replays.
"""
from __future__ import annotations

import json
import re
import tempfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

from measure_chart_markup import chart_part

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/chart_data_labels.json'
READ = 'sr.HasDataLabels & "," & sr.HasLeaderLines'
CASES = [
    (kind, statements)
    for kind in ('xlColumnClustered', 'xlLine', 'xlPie', 'xlXYScatter', 'xlArea', 'xlDoughnut')
    for statements in (f's = {READ}', f'sr.HasDataLabels = True: s = {READ}',
                       f'sr.HasDataLabels = True: sr.HasDataLabels = False: s = {READ}')
] + [
    ('xlColumnClustered', 'sr.HasDataLabels = 5: s = sr.HasDataLabels'),
    ('xlPie', f'sr.HasLeaderLines = False: s = {READ}'),
    ('xlPie', f'sr.HasDataLabels = True: sr.HasLeaderLines = False: s = {READ}'),
    ('xlColumnClustered', f'sr.HasLeaderLines = True: s = {READ}'),
    ('xlColumnClustered', 'sr.HasDataLabels = True: ch.ChartType = xlLine'),
]


def _probe() -> str:
    lines = ['Private Function Tried(ByVal which As Long, ByVal folder As String) As String',
             '    Dim wb As Workbook, ws As Worksheet, ch As Object, sr As Object, s As String',
             '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set ws = wb.Worksheets(1)',
             '    Set ch = ws.ChartObjects.Add(100, 10, 300, 200).Chart',
             '    ws.Range("A1:A4").Value = Application.Transpose(Array("k", "a", "b", "c"))',
             '    ws.Range("B1:B4").Value = Application.Transpose(Array("v", 1, -2, 3))',
             '    ws.Range("C1:C4").Value = Application.Transpose(Array("w", 4, 0.5, 6))',
             '    ch.SetSourceData ws.Range("A1:C4")', '    Set sr = ch.SeriesCollection(1)',
             '    On Error GoTo Failed', '    Select Case which']
    for index, (kind, statements) in enumerate(CASES):
        lines += [f'    Case {index}', f'        ch.ChartType = {kind}', f'        {statements}']
    lines += ['    End Select', '    Tried = "ok " & s', 'Saved:',
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
        found = re.findall(r'<c:ser>.*?</c:ser>|<c:dLbls>.*?</c:dLbls>', part)
        print(case, '->', answer, '||', ' '.join(found)[:400])


if __name__ == '__main__':
    main()
