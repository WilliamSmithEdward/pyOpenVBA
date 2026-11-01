"""Axis.Delete and Axis.CategoryNames: what a chart reads after each, the errors they raise, and the chart part Excel
saves.

Each case makes a workbook, adds a chart through ChartObjects.Add, fills categories in A2:A4 and two series in B and
C, gives them to the chart with SetSourceData (unless the case builds its own series), sets its type and runs the
case's statements, then saves it. The chart part is recorded as scripts/measure_chart_markup.py records it.

    python scripts/measure_chart_axis_names.py

writes tests/fixtures/chart_axis_names.json, which tests/test_excel_chart_axis_names.py replays.
"""
from __future__ import annotations

import json
import tempfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

from measure_chart_markup import chart_part

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/chart_axis_names.json'
READ = ('"|" & ch.HasAxis(1) & "," & ch.HasAxis(2) & "," & ch.SeriesCollection.Count & "," & '
        'ch.SeriesCollection(1).Formula')
NAMES = 's = s & "|" & TypeName(v) & ":" & LBound(v) & ":" & UBound(v) & ":" & Join(v, ";")'
#: (chart type, whether SetSourceData runs, statements)
CASES = [
    ('xlColumnClustered', True, 'v = ch.Axes(1).CategoryNames: ' + NAMES),
    ('xlColumnClustered', True, 'ch.Axes(1).CategoryNames = Array("x", "y", "z"): v = ch.Axes(1).CategoryNames: '
     + NAMES),
    ('xlColumnClustered', True, 'ch.Axes(1).CategoryNames = ws.Range("C2:C4"): v = ch.Axes(1).CategoryNames: '
     + NAMES),
    ('xlColumnClustered', True, 'ch.Axes(1).CategoryNames = Array(1, 2, 3)'),
    ('xlColumnClustered', True, 'ch.Axes(1).CategoryNames = Array("only")'),
    ('xlColumnClustered', True, 'ch.Axes(1).CategoryNames = "abc"'),
    ('xlLine', True, 'ch.Axes(1).CategoryNames = Array("p", "q", "r", "s")'),
    ('xlColumnClustered', True, 'v = ch.Axes(2).CategoryNames: ' + NAMES),
    ('xlXYScatter', True, 'v = ch.Axes(1).CategoryNames: ' + NAMES),
    ('xlColumnClustered', False, 'ch.SeriesCollection.NewSeries.Values = Array(1, 2): '
     'v = ch.Axes(1).CategoryNames: ' + NAMES),
    ('xlColumnClustered', True, 'ch.Axes(1).Delete'),
    ('xlColumnClustered', True, 'ch.Axes(2).Delete'),
    ('xlColumnClustered', True, 'ch.Axes(1).Delete: ch.HasAxis(1) = True'),
    ('xlXYScatter', True, 'ch.Axes(1).Delete'),
    ('xlColumnClustered', True, 'Set a = ch.Axes(1): a.Delete: s = s & "|" & a.MajorTickMark'),
    ('xlColumnClustered', True, 'ch.SeriesCollection(1).XValues = ws.Range("C2:C4")'),
    ('xlColumnClustered', True, 'ch.SeriesCollection(1).Values = ws.Range("C2:C4")'),
]


def _probe() -> str:
    lines = ['Private Function Tried(ByVal which As Long, ByVal folder As String) As String',
             '    Dim wb As Workbook, ws As Worksheet, ch As Object, a As Object, s As String, v As Variant',
             '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set ws = wb.Worksheets(1)',
             '    Set ch = ws.ChartObjects.Add(100, 10, 300, 200).Chart',
             '    ws.Range("A1:A4").Value = Application.Transpose(Array("k", "a", "b", "c"))',
             '    ws.Range("B1:B4").Value = Application.Transpose(Array("v", 1, 2, 3))',
             '    ws.Range("C1:C4").Value = Application.Transpose(Array("w", 4, 0.5, 6))',
             '    On Error GoTo Failed', '    Select Case which']
    for index, (kind, source, statements) in enumerate(CASES):
        lines.append(f'    Case {index}')
        if source:
            lines.append('        ch.SetSourceData ws.Range("A1:C4")')
        lines += [f'        ch.ChartType = {kind}: s = {READ}', f'        {statements}', f'        s = s & {READ}']
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
                        for (kind, _, statements), answer, part in zip(CASES, answers, parts, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for case, answer, _ in record['cases']:
        print(case[:90], '->', answer)


if __name__ == '__main__':
    main()
