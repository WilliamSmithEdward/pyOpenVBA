"""Axis.MinorUnit, MinorUnitIsAuto, ScaleType, LogBase, AxisBetweenCategories, TickLabelSpacing, TickLabelSpacingIsAuto,
TickMarkSpacing, DisplayUnit, DisplayUnitCustom and HasDisplayUnitLabel: what an axis reads before and after each
change, the errors bad values raise, and the chart part Excel saves.

Each case makes a workbook, adds a chart through ChartObjects.Add, fills categories in A2:A4 and two series in B and
C, gives them to the chart with SetSourceData, sets its type and runs the case's statements on the axis it names,
then saves it. Members Excel answers only for one kind of axis are read where the case asks. The chart part is
recorded as scripts/measure_chart_markup.py records it.

    python scripts/measure_chart_axis_scale.py

writes tests/fixtures/chart_axis_scale.json, which tests/test_excel_chart_axis_scale.py replays.
"""
from __future__ import annotations

import json
import tempfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

from measure_chart_markup import chart_part

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/chart_axis_scale.json'
VALUE = ('"|" & a.MinorUnitIsAuto & "," & a.ScaleType & "," & a.DisplayUnit & "," & a.HasDisplayUnitLabel & "," & '
         'a.MajorUnitIsAuto')
CATEGORY = ('"|" & a.AxisBetweenCategories & "," & a.TickLabelSpacing & "," & a.TickLabelSpacingIsAuto & "," & '
            'a.TickMarkSpacing')
#: (chart type, axis, read, statements)
CASES = [
    ('xlColumnClustered', 2, 'VALUE', ''),
    ('xlColumnClustered', 1, 'CATEGORY', ''),
    ('xlXYScatter', 1, 'VALUE', ''),
    ('xlColumnClustered', 2, 'VALUE', 'a.MinorUnit = 0.25'),
    ('xlColumnClustered', 2, 'VALUE', 'a.MinorUnit = 0.25: a.MinorUnitIsAuto = True'),
    ('xlColumnClustered', 2, 'VALUE', 'a.MinorUnit = 0'),
    ('xlColumnClustered', 2, 'VALUE', 'a.MinorUnit = -1'),
    ('xlColumnClustered', 2, 'VALUE', 'a.MinorUnit = 0.5: s = s & "," & a.MinorUnit'),
    ('xlColumnClustered', 2, 'VALUE', 'a.ScaleType = xlScaleLogarithmic'),
    ('xlColumnClustered', 2, 'VALUE', 'a.ScaleType = xlScaleLogarithmic: s = s & "," & a.LogBase'),
    ('xlColumnClustered', 2, 'VALUE', 'a.ScaleType = xlScaleLogarithmic: a.LogBase = 2: s = s & "," & a.LogBase'),
    ('xlColumnClustered', 2, 'VALUE', 'a.ScaleType = xlScaleLogarithmic: a.ScaleType = xlScaleLinear'),
    ('xlColumnClustered', 2, 'VALUE', 'a.LogBase = 5: s = s & "," & a.LogBase'),
    ('xlColumnClustered', 2, 'VALUE', 'a.ScaleType = xlScaleLogarithmic: a.LogBase = 1'),
    ('xlColumnClustered', 2, 'VALUE', 'a.ScaleType = xlScaleLogarithmic: a.LogBase = 1000'),
    ('xlColumnClustered', 2, 'VALUE', 'a.ScaleType = 7'),
    ('xlColumnClustered', 2, 'VALUE', 'a.DisplayUnit = xlThousands'),
    ('xlColumnClustered', 2, 'VALUE', 'a.DisplayUnit = xlMillions: a.HasDisplayUnitLabel = False'),
    ('xlColumnClustered', 2, 'VALUE', 'a.DisplayUnit = xlCustom: s = s & "," & a.DisplayUnitCustom'),
    ('xlColumnClustered', 2, 'VALUE', 'a.DisplayUnitCustom = 250: s = s & "," & a.DisplayUnitCustom'),
    ('xlColumnClustered', 2, 'VALUE', 'a.DisplayUnit = xlHundreds: a.DisplayUnit = xlNone'),
    ('xlColumnClustered', 2, 'VALUE', 'a.DisplayUnit = 99'),
    ('xlColumnClustered', 2, 'VALUE', 'a.HasDisplayUnitLabel = True'),
    ('xlColumnClustered', 1, 'CATEGORY', 'a.AxisBetweenCategories = False'),
    ('xlLine', 1, 'CATEGORY', 'a.AxisBetweenCategories = False'),
    ('xlArea', 1, 'CATEGORY', ''),
    ('xlArea', 1, 'CATEGORY', 'a.AxisBetweenCategories = True'),
    ('xlColumnClustered', 1, 'CATEGORY', 'a.TickLabelSpacing = 2'),
    ('xlColumnClustered', 1, 'CATEGORY', 'a.TickLabelSpacing = 2: a.TickLabelSpacingIsAuto = True'),
    ('xlColumnClustered', 1, 'CATEGORY', 'a.TickMarkSpacing = 3'),
    ('xlColumnClustered', 1, 'CATEGORY', 'a.TickLabelSpacing = 0'),
    ('xlColumnClustered', 1, 'CATEGORY', 'a.TickMarkSpacing = 32000'),
    ('xlColumnClustered', 1, 'CATEGORY', 'a.TickMarkSpacing = 32001'),
    ('xlColumnClustered', 1, 'CATEGORY', 'a.TickLabelSpacing = 2.5'),
    ('xlColumnClustered', 1, 'VALUE', ''),
    ('xlColumnClustered', 2, 'CATEGORY', ''),
    ('xlColumnClustered', 1, 'CATEGORY', 'a.MinorUnit = 1'),
]


def _probe() -> str:
    lines = ['Private Function Tried(ByVal which As Long, ByVal folder As String) As String',
             '    Dim wb As Workbook, ws As Worksheet, ch As Object, a As Object, s As String',
             '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set ws = wb.Worksheets(1)',
             '    Set ch = ws.ChartObjects.Add(100, 10, 300, 200).Chart',
             '    ws.Range("A1:A4").Value = Application.Transpose(Array("k", "a", "b", "c"))',
             '    ws.Range("B1:B4").Value = Application.Transpose(Array("v", 1, 2, 3))',
             '    ws.Range("C1:C4").Value = Application.Transpose(Array("w", 4, 0.5, 6))',
             '    ch.SetSourceData ws.Range("A1:C4")', '    On Error GoTo Failed', '    Select Case which']
    for index, (kind, axis, read, statements) in enumerate(CASES):
        reader = VALUE if read == 'VALUE' else CATEGORY
        lines += [f'    Case {index}', f'        ch.ChartType = {kind}: Set a = ch.Axes({axis}): s = {reader}']
        if statements:
            lines.append(f'        {statements}')
        lines.append(f'        s = s & {reader}')
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
                        for (kind, axis, _, statements), answer, part in zip(CASES, answers, parts, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for case, answer, _ in record['cases']:
        print(case, '->', answer)


if __name__ == '__main__':
    main()
