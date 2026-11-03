"""More of a chart's members: a trendline's Forward2, Backward2 and DataLabel, DataLabels.Format, and Axis.MinorUnitScale. What
each reads, and the chart part Excel saves.

Each case makes a workbook with numbers in A2:A5 and B2:B5, a chart of them of the type it names, with a trendline on its
series where the case wants one. It runs the case's statements, which set r to what they want recorded, and saves the
workbook. The chart part is recorded as scripts/measure_chart_markup.py records it.

    python scripts/measure_trendline_members.py

writes tests/fixtures/trendline_members.json, which tests/test_excel_trendline_members.py replays.
"""
from __future__ import annotations

import json
import tempfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

from measure_chart_markup import chart_part

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/trendline_members.json'
S1 = 'ch.SeriesCollection(1)'
T = f'{S1}.Trendlines(1)'
ADD = f'{S1}.Trendlines.Add: '
LABELS = f'{S1}.HasDataLabels = True: '
L = f'{S1}.DataLabels'
CASES = [
    # How far it runs.
    ('xlXYScatter', f'{ADD}r = {T}.Forward2 & " " & {T}.Backward2 & " " & {T}.Forward & " " & {T}.Backward'),
    *[('xlXYScatter', f'{ADD}{T}.Forward2 = {value}: r = {T}.Forward2 & " " & {T}.Forward') for value in (1, 0.5, 2.25, 0, -1)],
    *[('xlXYScatter', f'{ADD}{T}.Backward2 = {value}: r = {T}.Backward2 & " " & {T}.Backward') for value in (1, 0.5, 2.25, 0, -1)],
    *[('xlXYScatter', f'{ADD}{T}.Forward = {value}: r = {T}.Forward2 & " " & {T}.Forward') for value in (1, 2)],
    *[('xlXYScatter', f'{ADD}{T}.Backward = {value}: r = {T}.Backward2 & " " & {T}.Backward') for value in (1, 2)],
    *[('xlColumnClustered', f'{ADD}{T}.Forward2 = {value}: r = {T}.Forward2') for value in (0.5, 1)], *[('xlColumnClustered', f'{ADD}{T}.Backward2 = {value}: r = {T}.Backward2') for value in (0.5, 1)],
    ('xlLine', f'{ADD}{T}.Forward2 = 1.5: r = {T}.Forward2'), ('xlXYScatter', f'{ADD}{T}.Forward2 = "x"'),
    # Its label.
    ('xlXYScatter', f'{ADD}r = TypeName({T}.DataLabel)'), ('xlXYScatter', f'{ADD}{T}.DisplayEquation = True: r = TypeName({T}.DataLabel)'), ('xlXYScatter', f'{ADD}{T}.DisplayEquation = True: r = {T}.DataLabel.Text'), ('xlXYScatter', f'{ADD}{T}.DisplayRSquared = True: r = {T}.DataLabel.Text'), ('xlXYScatter', f'{ADD}{T}.DisplayEquation = True: {T}.DisplayRSquared = True: r = {T}.DataLabel.Text'),
    ('xlXYScatter', f'{ADD}{T}.DisplayEquation = True: {T}.DataLabel.Text = "mine": r = {T}.DataLabel.Text'), ('xlXYScatter', f'{ADD}{T}.DisplayEquation = True: r = {T}.DataLabel.NumberFormat'), ('xlXYScatter', f'{ADD}{T}.DisplayEquation = True: {T}.DataLabel.NumberFormat = "0.00": r = {T}.DataLabel.Text'), ('xlXYScatter', f'{ADD}{T}.DisplayEquation = True: r = {T}.DataLabel.Font.Size'), ('xlXYScatter', f'{ADD}{T}.DisplayEquation = True: {T}.DataLabel.Font.Size = 14: r = {T}.DataLabel.Font.Size'), ('xlXYScatter', f'{ADD}{T}.DisplayEquation = True: r = {T}.DataLabel.Name'), ('xlXYScatter', f'{ADD}{T}.DisplayEquation = True: {T}.DataLabel.Delete: r = {T}.DisplayEquation'), ('xlXYScatter', f'{ADD}{T}.DisplayEquation = True: r = {T}.DataLabel.Format.Fill.Visible'), ('xlXYScatter', f'{ADD}{T}.DisplayEquation = True: {T}.DataLabel.Format.Fill.ForeColor.RGB = 255: r = {T}.DataLabel.Format.Fill.ForeColor.RGB'),
    # The labels' Format.
    ('xlColumnClustered', f'{LABELS}r = TypeName({L}.Format)'), ('xlColumnClustered', f'{LABELS}r = {L}.Format.Fill.Visible'), ('xlColumnClustered', f'{LABELS}r = {L}.Format.Line.Visible'), ('xlColumnClustered', f'{LABELS}r = {L}.Format.Fill.ForeColor.RGB'), ('xlColumnClustered', f'{LABELS}{L}.Format.Fill.ForeColor.RGB = 255: r = {L}.Format.Fill.ForeColor.RGB'), ('xlColumnClustered', f'{LABELS}{L}.Format.Line.ForeColor.RGB = 255: r = {L}.Format.Line.ForeColor.RGB'), ('xlColumnClustered', f'{LABELS}{L}.Format.Line.Weight = 2: r = {L}.Format.Line.Weight'), ('xlColumnClustered', f'{LABELS}{L}.Format.Fill.Visible = msoFalse: r = {L}.Format.Fill.Visible'), ('xlColumnClustered', f'{LABELS}{L}.Format.Line.Visible = msoFalse: r = {L}.Format.Line.Visible'), ('xlColumnClustered', f'{LABELS}{L}.Format.Fill.ForeColor.RGB = 255: r = {S1}.Points(1).DataLabel.Format.Fill.ForeColor.RGB'), ('xlColumnClustered', f'{LABELS}{L}.Format.Fill.ForeColor.RGB = 255: {S1}.Points(1).DataLabel.Format.Fill.ForeColor.RGB = 65280: r = {L}.Format.Fill.ForeColor.RGB'), ('xlColumnClustered', f'{LABELS}r = TypeName({L}.Format.Parent)'), ('xlColumnClustered', f'r = TypeName({L}.Format)'), ('xlColumnClustered', f'{LABELS}r = {L}.Propagate'),
    # A date axis's minor unit.
    ('xlColumnClustered', 'r = ch.Axes(xlCategory).MinorUnitScale'), ('xlColumnClustered', 'r = ch.Axes(xlValue).MinorUnitScale'), ('xlColumnClustered', 'ch.Axes(xlCategory).CategoryType = xlTimeScale: r = ch.Axes(xlCategory).MinorUnitScale'), ('xlColumnClustered', 'ch.Axes(xlCategory).CategoryType = xlTimeScale: ch.Axes(xlCategory).MinorUnitScale = xlMonths: r = ch.Axes(xlCategory).MinorUnitScale'), ('xlColumnClustered', 'ch.Axes(xlCategory).CategoryType = xlTimeScale: ch.Axes(xlCategory).MinorUnitScale = xlYears: r = ch.Axes(xlCategory).MinorUnitScale'), ('xlColumnClustered', 'ch.Axes(xlCategory).CategoryType = xlTimeScale: ch.Axes(xlCategory).MinorUnitScale = xlDays: r = ch.Axes(xlCategory).MinorUnitScale'), ('xlColumnClustered', 'ch.Axes(xlCategory).CategoryType = xlTimeScale: ch.Axes(xlCategory).MinorUnitScale = 9'), ('xlColumnClustered', 'ch.Axes(xlCategory).MinorUnitScale = xlMonths: r = ch.Axes(xlCategory).MinorUnitScale'), ('xlColumnClustered', 'ch.Axes(xlValue).MinorUnitScale = xlMonths'), ('xlXYScatter', 'r = ch.Axes(xlCategory).MinorUnitScale'),
]


def _probe() -> str:
    lines = ['Private Function Tried(ByVal which As Long, ByVal folder As String) As String',
             '    Dim wb As Workbook, ws As Object, ch As Object, s As String, v As String, r As Variant', '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set ws = wb.Worksheets(1)',
             '    ws.Range("A1:A5").Value = Application.Transpose(Array("x", 1, 2, 4, 8))', '    ws.Range("B1:B5").Value = Application.Transpose(Array("v", 3, 5, 4, 9))',
             '    Set ch = ws.ChartObjects.Add(300, 100, 300, 200).Chart', '    ch.SetSourceData ws.Range("A1:B5"), xlColumns', '    On Error GoTo Failed', '    Select Case which']
    for index, (kind, statement) in enumerate(CASES):
        lines += [f'    Case {index}', f'        ch.ChartType = {kind}', f'        {statement}']
    lines += ['    End Select', '    s = "ok "', 'Answered:', '    On Error Resume Next', '    Err.Clear', '    v = "": v = CStr(r)', '    If Err.Number Then v = "E" & Err.Number', '    s = s & "|" & v',
              '    On Error GoTo 0', '    Tried = s', '    wb.SaveAs folder & "\\c" & which & ".xlsx", xlOpenXMLWorkbook', '    wb.Close False', '    Exit Function',
              'Failed:', '    s = "Err " & Err.Number & " "', '    Resume Answered', 'End Function', '',
              'Public Function Probe(ByVal folder As String, ByVal first As Long, ByVal last As Long) As String', '    Dim i As Long, s As String',
              '    Application.DisplayAlerts = False', '    For i = first To last', '        s = s & "@@@" & Tried(i, folder)', '    Next', '    Application.DisplayAlerts = True',
              '    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def main() -> None:
    source = _probe()
    answers: list[str] = []
    with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        for first in range(0, len(CASES), 30):
            last = min(first + 29, len(CASES) - 1)
            result = excel.run_vba(source, 'Probe', (folder, first, last), timeout=280.0)
            assert result.ok and not result.dialogs, result
            answers += str(result.value).split('@@@')
        parts = [chart_part(Path(folder) / f'c{index}.xlsx') for index in range(len(CASES))]
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source,
              'cases': [[f'{kind}: {statement}', answer, part] for (kind, statement), answer, part in zip(CASES, answers, parts, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for case, answer, _ in record['cases']:
        print(case, '=>', answer)


if __name__ == '__main__':
    main()
