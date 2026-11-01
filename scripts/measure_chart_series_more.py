"""Series.InvertIfNegative and Series.Explosion, and Chart.HasDataTable and Chart.Name: what they read before and
after each change, the errors, and the chart part Excel saves.

Each case makes a workbook, adds a chart through ChartObjects.Add, fills categories in A2:A4 and two series in B and
C, gives them to the chart with SetSourceData, sets the case's chart type and runs the case's statements with sr as
the first series, then saves it. The chart part is recorded as scripts/measure_chart_markup.py records it.

    python scripts/measure_chart_series_more.py

writes tests/fixtures/chart_series_more.json, which tests/test_excel_chart_series_more.py replays.
"""
from __future__ import annotations

import json
import re
import tempfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

from measure_chart_markup import chart_part

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/chart_series_more.json'
CASES = [
    ('xlColumnClustered', 's = sr.InvertIfNegative'),
    ('xlColumnClustered', 'sr.InvertIfNegative = True'),
    ('xlColumnClustered', 'sr.InvertIfNegative = True: sr.InvertIfNegative = False'),
    ('xlBarClustered', 'sr.InvertIfNegative = True'),
    ('xlLine', 's = sr.InvertIfNegative'),
    ('xlLine', 'sr.InvertIfNegative = True'),
    ('xlPie', 's = sr.Explosion'),
    ('xlPie', 'sr.Explosion = 25'),
    ('xlPie', 'sr.Explosion = 0'),
    ('xlPie', 'sr.Explosion = 400'),
    ('xlPie', 'sr.Explosion = 401'),
    ('xlPie', 'sr.Explosion = -1'),
    ('xlDoughnut', 'sr.Explosion = 10'),
    ('xlColumnClustered', 's = sr.Explosion'),
    ('xlColumnClustered', 'sr.Explosion = 10'),
    ('xlColumnClustered', 's = ch.HasDataTable'),
    ('xlColumnClustered', 'ch.HasDataTable = True'),
    ('xlColumnClustered', 'ch.HasDataTable = True: ch.HasDataTable = False'),
    ('xlLine', 'ch.HasDataTable = True'),
    ('xlPie', 'ch.HasDataTable = True'),
    ('xlColumnClustered', 's = ch.Name'),
    ('xlColumnClustered', 'ch.Parent.Name = "Sales": s = ch.Name'),
    ('xlColumnClustered', 'ch.Name = "X"'),
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
        found = re.findall(r'<c:invertIfNegative[^>]*>|<c:explosion[^>]*>|<c:dTable>.*?</c:dTable>|'
                           r'<c:ext uri="\{6F2FDCE9[^}]*\}".*?</c:ext>', part)
        print(case, '->', answer, '||', ' '.join(found)[:400])


if __name__ == '__main__':
    main()
