"""Axis.MajorTickMark, MinorTickMark, TickLabelPosition, HasMajorGridlines, HasMinorGridlines, Crosses and CrossesAt:
what an axis reads before and after each change, the errors bad values raise, and the chart part Excel saves.

Each case makes a workbook, adds a chart through ChartObjects.Add, fills categories in A2:A4 and two series in B and
C and gives them to the chart with SetSourceData, sets its type and runs the case's statements on the axis it names, then saves it. CrossesAt is
read only where the case asks, as Excel works it out from the drawn scale when the crossing is xlMaximum or xlMinimum.
The chart part is recorded as scripts/measure_chart_markup.py records it.

    python scripts/measure_chart_axis_lines.py

writes tests/fixtures/chart_axis_lines.json, which tests/test_excel_chart_axis_lines.py replays.
"""
from __future__ import annotations

import json
import tempfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

from measure_chart_markup import chart_part

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/chart_axis_lines.json'
READ = ('"|" & a.MajorTickMark & "," & a.MinorTickMark & "," & a.TickLabelPosition & "," & a.HasMajorGridlines & '
        '"," & a.HasMinorGridlines & "," & a.Crosses')
#: (chart type, axis, statements, whether CrossesAt is read after them)
CASES = [
    ('xlColumnClustered', 1, '', True),
    ('xlColumnClustered', 2, '', True),
    ('xlXYScatter', 1, '', True),
    ('xlBarClustered', 2, '', True),
    ('xlColumnClustered', 1, 'a.MajorTickMark = xlTickMarkInside', False),
    ('xlColumnClustered', 2, 'a.MajorTickMark = xlTickMarkCross', False),
    ('xlColumnClustered', 2, 'a.MajorTickMark = xlTickMarkNone', False),
    ('xlColumnClustered', 2, 'a.MinorTickMark = xlTickMarkOutside', False),
    ('xlColumnClustered', 1, 'a.MinorTickMark = xlTickMarkInside: a.MajorTickMark = xlTickMarkOutside', False),
    ('xlColumnClustered', 1, 'a.MajorTickMark = 9', False),
    ('xlColumnClustered', 1, 'a.MajorTickMark = "2"', False),
    ('xlColumnClustered', 1, 'a.TickLabelPosition = xlTickLabelPositionLow', False),
    ('xlColumnClustered', 2, 'a.TickLabelPosition = xlTickLabelPositionHigh', False),
    ('xlColumnClustered', 2, 'a.TickLabelPosition = xlTickLabelPositionNone', False),
    ('xlColumnClustered', 2, 'a.TickLabelPosition = 1', False),
    ('xlColumnClustered', 2, 'a.HasMajorGridlines = False', False),
    ('xlColumnClustered', 1, 'a.HasMajorGridlines = True', False),
    ('xlColumnClustered', 2, 'a.HasMinorGridlines = True', False),
    ('xlColumnClustered', 1, 'a.HasMinorGridlines = True: a.HasMajorGridlines = True', False),
    ('xlXYScatter', 1, 'a.HasMajorGridlines = True: a.HasMinorGridlines = True', False),
    ('xlColumnClustered', 2, 'a.HasMajorGridlines = "x"', False),
    ('xlColumnClustered', 1, 'a.Crosses = xlMaximum', False),
    ('xlColumnClustered', 2, 'a.Crosses = xlMaximum', False),
    ('xlColumnClustered', 2, 'a.Crosses = xlMinimum', False),
    ('xlColumnClustered', 2, 'a.Crosses = xlMinimum: a.Crosses = xlAxisCrossesAutomatic', True),
    ('xlColumnClustered', 2, 'a.CrossesAt = 1.5', True),
    ('xlColumnClustered', 1, 'a.CrossesAt = 2', True),
    ('xlColumnClustered', 2, 'a.Crosses = xlAxisCrossesCustom', True),
    ('xlColumnClustered', 2, 'a.CrossesAt = 1.5: a.Crosses = xlAxisCrossesCustom', True),
    ('xlColumnClustered', 2, 'a.CrossesAt = 1.5: a.Crosses = xlAxisCrossesAutomatic', True),
    ('xlColumnClustered', 2, 'a.Crosses = 7', False),
    ('xlXYScatter', 1, 'a.CrossesAt = 2.5', True),
    ('xlXYScatter', 2, 'a.Crosses = xlMaximum', False),
    ('xlColumnClustered', 1, 'a.Crosses = xlMaximum: a.ReversePlotOrder = True', False),
    ('xlColumnClustered', 2, 'a.Crosses = xlMaximum: ch.Axes(1).ReversePlotOrder = True', False),
    ('xlColumnClustered', 2,
     'a.MajorTickMark = xlTickMarkInside: a.HasMinorGridlines = True: a.CrossesAt = 2: ch.ChartType = xlLine', True),
]


def _probe() -> str:
    lines = ['Private Function Tried(ByVal which As Long, ByVal folder As String) As String',
             '    Dim wb As Workbook, ws As Worksheet, ch As Object, a As Object, s As String',
             '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set ws = wb.Worksheets(1)',
             '    Set ch = ws.ChartObjects.Add(100, 10, 300, 200).Chart',
             '    ws.Range("A1:A4").Value = Application.Transpose(Array("k", "a", "b", "c"))',
             '    ws.Range("B1:B4").Value = Application.Transpose(Array("v", 1, 2, 3))',
             '    ws.Range("C1:C4").Value = Application.Transpose(Array("w", 4, 0.5, 6))',
             '    ch.SetSourceData ws.Range("A1:C4")',
             '    On Error GoTo Failed', '    Select Case which']
    for index, (kind, axis, statements, crossing) in enumerate(CASES):
        lines += [f'    Case {index}', f'        ch.ChartType = {kind}: Set a = ch.Axes({axis}): s = {READ}']
        if statements:
            lines.append(f'        {statements}')
        lines.append(f'        s = s & {READ}' + (' & "," & a.CrossesAt' if crossing else ''))
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
              'cases': [[f'{kind} {axis}: {statements}', answer, part]
                        for (kind, axis, statements, _), answer, part in zip(CASES, answers, parts, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for case, answer, _ in record['cases']:
        print(case, '->', answer)


if __name__ == '__main__':
    main()
