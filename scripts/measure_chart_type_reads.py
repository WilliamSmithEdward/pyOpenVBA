"""Chart members that read a fixed value for each chart type the model makes: the 3-D view settings, ChartStyle,
ChartColor and the label levels, and the line objects of the chart's first group.

Each case makes a workbook, adds a chart through ChartObjects.Add, fills categories in A2:A4 and two series in B and
C, gives them to the chart with SetSourceData, sets the case's type and runs its statements with cg as the first
group, then reads READS, each as its value (TypeName for an object) or "E" and the error number.

    python scripts/measure_chart_type_reads.py

writes tests/fixtures/chart_type_reads.json, which tests/test_excel_chart_type_reads.py replays.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/chart_type_reads.json'
#: The chart types the model makes (_charts.CHART_TYPES).
TYPES = [51, 52, 53, 57, 58, 59, 4, 63, 64, 65, 66, 67, 5, -4120, -4169, 1, 76, 77]
CASES = [(kind, '') for kind in TYPES] + [
    (4, 'cg.HasDropLines = True'), (4, 'cg.HasHiLoLines = True'), (4, 'cg.HasUpDownBars = True'),
    (52, 'cg.HasSeriesLines = True'), (4, 'cg.HasDropLines = True: cg.DropLines.Delete'),
    (4, 'cg.HasUpDownBars = True: cg.UpBars.Delete'),
]
READS = [
    'ch.Elevation', 'ch.Rotation', 'ch.Perspective', 'ch.RightAngleAxes', 'ch.DepthPercent', 'ch.HeightPercent',
    'ch.GapDepth', 'ch.AutoScaling', 'ch.BarShape', 'ch.ChartStyle', 'ch.ChartColor', 'ch.CategoryLabelLevel',
    'ch.SeriesNameLevel',
    'cg.DropLines.Name', 'TypeName(cg.DropLines.Parent)', 'cg.HiLoLines.Name', 'cg.UpBars.Name', 'cg.DownBars.Name',
    'TypeName(cg.UpBars.Parent)', 'cg.SeriesLines.Name', 'TypeName(cg.SeriesLines.Application)',
    'cg.HasDropLines', 'cg.HasUpDownBars',
]


def _probe() -> str:
    lines = ['Private Function Tried(ByVal which As Long) As String',
             '    Dim wb As Workbook, ws As Worksheet, ch As Object, cg As Object, s As String',
             '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set ws = wb.Worksheets(1)',
             '    Set ch = ws.ChartObjects.Add(100, 10, 300, 200).Chart',
             '    ws.Range("A1:A4").Value = Application.Transpose(Array("k", "a", "b", "c"))',
             '    ws.Range("B1:B4").Value = Application.Transpose(Array("v", 1, 2, 3))',
             '    ws.Range("C1:C4").Value = Application.Transpose(Array("w", 4, 5, 6))',
             '    ch.SetSourceData ws.Range("A1:C4")', '    On Error GoTo Failed', '    Select Case which']
    for index, (kind, statements) in enumerate(CASES):
        lines += [f'    Case {index}', f'        ch.ChartType = {kind}', '        Set cg = ch.ChartGroups(1)']
        lines += [f'        {statements}'] if statements else []
    lines += ['    End Select', '    On Error Resume Next']
    for read in READS:
        lines += ['    Err.Clear', f'    s = s & "," & {read}', '    If Err.Number Then s = s & "E" & Err.Number']
    lines += ['    On Error GoTo 0', '    Tried = "ok " & Mid(s, 2)', 'Closed:', '    wb.Close False',
              '    Exit Function', 'Failed:', '    Tried = "Err " & Err.Number', '    Resume Closed', 'End Function',
              '', 'Public Function Probe() As String', '    Dim i As Long, s As String',
              f'    For i = 0 To {len(CASES) - 1}', '        s = s & "@@@" & Tried(i)', '    Next',
              '    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def main() -> None:
    source = _probe()
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(source, 'Probe', timeout=280.0)
        assert result.ok and not result.dialogs, result
        answers = str(result.value).split('@@@')
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'reads': READS, 'source': source,
              'cases': [[f'{kind}: {statements}', answer] for (kind, statements), answer in zip(CASES, answers,
                                                                                               strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for case, answer in record['cases']:
        print(case, '->', answer)


if __name__ == '__main__':
    main()
