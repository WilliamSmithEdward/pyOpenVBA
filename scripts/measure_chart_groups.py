"""Chart.ChartGroups and a ChartGroup's GapWidth, Overlap, VaryByCategories, FirstSliceAngle and DoughnutHoleSize:
what they read before and after each change, the errors, and the chart part Excel saves.

Each case makes a workbook, adds a chart through ChartObjects.Add, fills categories in A2:A4 and two series in B and
C, gives them to the chart with SetSourceData, sets the case's chart type and runs the case's statements with sr as
the first series, then saves it. The chart part is recorded as scripts/measure_chart_markup.py records it.

    python scripts/measure_chart_groups.py

writes tests/fixtures/chart_groups.json, which tests/test_excel_chart_groups.py replays.
"""
from __future__ import annotations

import json
import re
import tempfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

from measure_chart_markup import chart_part

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/chart_groups.json'
G = 'ch.ChartGroups(1)'
CASES = [
    *((kind, '') for kind in ('xlColumnClustered', 'xlBarClustered', 'xlLine', 'xlPie', 'xlDoughnut', 'xlArea',
                              'xlXYScatter')),
    ('xlColumnClustered', f'{G}.GapWidth = 80'),
    ('xlColumnClustered', f'{G}.GapWidth = 0'),
    ('xlColumnClustered', f'{G}.GapWidth = 500'),
    ('xlColumnClustered', f'{G}.GapWidth = 501'),
    ('xlColumnClustered', f'{G}.Overlap = 50'),
    ('xlColumnClustered', f'{G}.Overlap = -100'),
    ('xlColumnClustered', f'{G}.Overlap = 101'),
    ('xlBarClustered', f'{G}.Overlap = -30: {G}.GapWidth = 200'),
    ('xlColumnClustered', f'{G}.VaryByCategories = True'),
    ('xlLine', f'{G}.VaryByCategories = True'),
    ('xlPie', f'{G}.VaryByCategories = False'),
    ('xlPie', f'{G}.FirstSliceAngle = 90'),
    ('xlPie', f'{G}.FirstSliceAngle = 360'),
    ('xlPie', f'{G}.FirstSliceAngle = 361'),
    ('xlDoughnut', f'{G}.DoughnutHoleSize = 75'),
    ('xlDoughnut', f'{G}.DoughnutHoleSize = 10'),
    ('xlDoughnut', f'{G}.DoughnutHoleSize = 91'),
    ('xlDoughnut', f'{G}.FirstSliceAngle = 45'),
    ('xlLine', f'{G}.GapWidth = 80'),
    ('xlPie', f'{G}.DoughnutHoleSize = 60'),
    ('xlColumnClustered', f'{G}.GapWidth = 80: ch.ChartType = xlBarClustered'),
    ('xlColumnClustered', 's = TypeName(ch.ChartGroups(2))'),
]
NAMES = ('GapWidth', 'Overlap', 'VaryByCategories', 'FirstSliceAngle', 'DoughnutHoleSize')


def _probe() -> str:
    lines = ['Private Function Safe(ByVal target As Object, ByVal name As String) As String',
             '    On Error GoTo Failed', '    Safe = CallByName(target, name, VbGet)', '    Exit Function', 'Failed:',
             '    Safe = "E" & Err.Number', 'End Function', '',
             'Private Function Tried(ByVal which As Long, ByVal folder As String) As String',
             '    Dim wb As Workbook, ws As Worksheet, ch As Object, sr As Object, s As String',
             '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set ws = wb.Worksheets(1)',
             '    Set ch = ws.ChartObjects.Add(100, 10, 300, 200).Chart',
             '    ws.Range("A1:A4").Value = Application.Transpose(Array("k", "a", "b", "c"))',
             '    ws.Range("B1:B4").Value = Application.Transpose(Array("v", 1, -2, 3))',
             '    ws.Range("C1:C4").Value = Application.Transpose(Array("w", 4, 0.5, 6))',
             '    ch.SetSourceData ws.Range("A1:C4")', '    Set sr = ch.SeriesCollection(1)',
             '    On Error GoTo Failed', '    Select Case which']
    for index, (kind, statements) in enumerate(CASES):
        lines += [f'    Case {index}', f'        ch.ChartType = {kind}'] + ([f'        {statements}'] if statements else [])
    lines += ['    End Select',
              '    s = s & "|" & ch.ChartGroups.Count & "," & ' + ' & "," & '.join(
                  f'Safe(ch.ChartGroups(1), "{name}")' for name in NAMES),
              '    Tried = "ok " & s', 'Saved:',
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
        found = re.findall(r'<c:gapWidth[^>]*>|<c:overlap[^>]*>|<c:varyColors[^>]*>|<c:firstSliceAng[^>]*>|<c:holeSize[^>]*>', part)
        print(case, '->', answer, '||', ' '.join(found)[:400])


if __name__ == '__main__':
    main()
