"""ApplyDataLabels on a chart, a series and a point: what each argument does, what it reads afterwards, and the chart part
Excel saves.

Each case makes a workbook with a chart of categories in A2:A4 and two series in B and C, of the chart type TYPES names,
calls ApplyDataLabels on the owner with the arguments, reads READS, each as its value or "E" and an error number, and
saves the workbook. The chart part is recorded as scripts/measure_chart_markup.py records it.

    python scripts/measure_apply_data_labels.py

writes tests/fixtures/apply_data_labels.json, which tests/test_excel_apply_data_labels.py replays.
"""
from __future__ import annotations

import json
import re
import tempfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

from measure_chart_markup import chart_part

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/apply_data_labels.json'
OWNERS = {'chart': 'ch', 'series': 'ch.SeriesCollection(1)', 'point': 'ch.SeriesCollection(1).Points(2)'}
NOBODY = 'nobody'
TYPES = {'column': 'xlColumnClustered', 'pie': 'xlPie', 'line': 'xlLine'}
ARGUMENTS = [
    '', 'xlDataLabelsShowValue', 'xlDataLabelsShowLabel', 'xlDataLabelsShowPercent', 'xlDataLabelsShowLabelAndPercent', 'xlDataLabelsShowBubbleSizes', 'xlDataLabelsShowNone', '99',
    'Type:=xlDataLabelsShowValue', 'LegendKey:=True', 'LegendKey:=False', 'AutoText:=True', 'AutoText:=False', 'HasLeaderLines:=True', 'HasLeaderLines:=False',
    'ShowSeriesName:=True', 'ShowCategoryName:=True', 'ShowValue:=True', 'ShowValue:=False', 'ShowPercentage:=True', 'ShowBubbleSize:=True', 'Separator:="; "',
    'ShowSeriesName:=True, ShowCategoryName:=True, ShowValue:=False', 'ShowSeriesName:=True, ShowCategoryName:=True, Separator:=" - "', 'xlDataLabelsShowLabel, ShowValue:=True',
    'xlDataLabelsShowValue, ShowValue:=False', 'xlDataLabelsShowNone, ShowValue:=True', 'ShowValue:=False, ShowCategoryName:=False', 'xlDataLabelsShowValue, True',
    'xlDataLabelsShowValue, False, True, False, False, False, True',
]
BEFORE = {'': '', 'dressed': 'ch.SeriesCollection(1).HasDataLabels = True: ch.SeriesCollection(1).DataLabels.Position = xlLabelPositionInsideEnd: ch.SeriesCollection(1).DataLabels.Separator = "/": ch.SeriesCollection(1).DataLabels.NumberFormat = "0.0": ch.SeriesCollection(1).DataLabels.Font.Bold = True',
          'alone': 'ch.SeriesCollection(1).Points(2).HasDataLabel = True', 'gone': 'ch.SeriesCollection(1).HasDataLabels = True: ch.SeriesCollection(1).Points(2).HasDataLabel = False',
          'on': 'ch.SeriesCollection(1).HasDataLabels = True', 'typed': 'ch.SeriesCollection(1).HasDataLabels = True: ch.SeriesCollection(1).Points(2).DataLabel.Text = "typed"',
          'cat': 'ch.SeriesCollection(1).HasDataLabels = True: ch.SeriesCollection(1).DataLabels.ShowCategoryName = True',
          'parts': 'ch.SeriesCollection(1).HasDataLabels = True: ch.SeriesCollection(1).DataLabels.ShowCategoryName = True: ch.SeriesCollection(1).DataLabels.ShowPercentage = True: ch.SeriesCollection(1).DataLabels.ShowValue = False',
          'all parts': 'ch.SeriesCollection(1).HasDataLabels = True: ch.SeriesCollection(1).DataLabels.ShowCategoryName = True: ch.SeriesCollection(1).DataLabels.ShowPercentage = True'}
CASES = [(owner, kind, '', argument) for owner in OWNERS for kind in TYPES for argument in ARGUMENTS]
CASES += [(owner, 'column', before, argument) for owner in OWNERS for before in ('on', 'typed', 'cat')
          for argument in ('', 'xlDataLabelsShowNone', 'xlDataLabelsShowLabel', 'ShowCategoryName:=True', 'ShowValue:=False')]
# How a pie joins a category and a percent, whichever way they came to be shown.
CASES += [(owner, 'pie', '', argument) for owner in OWNERS
          for argument in ('xlDataLabelsShowPercent, ShowCategoryName:=True', 'ShowCategoryName:=True, ShowPercentage:=True, ShowValue:=False',
                           'xlDataLabelsShowLabelAndPercent, ShowValue:=True', 'xlDataLabelsShowLabelAndPercent, Separator:="/"')]
CASES += [('nobody', 'pie', before, '') for before in ('parts', 'all parts')]
# What it leaves of labels already changed, and of a point's own label or the lack of one.
CASES += [(owner, 'column', before, argument) for owner in OWNERS for before in ('dressed', 'alone', 'gone')
          for argument in ('', 'xlDataLabelsShowLabel', 'xlDataLabelsShowNone')]
READS = [
    'ch.SeriesCollection(1).HasDataLabels', 'ch.SeriesCollection(2).HasDataLabels', 'ch.SeriesCollection(1).Points(1).HasDataLabel', 'ch.SeriesCollection(1).Points(2).HasDataLabel',
    'ch.SeriesCollection(1).Points(2).DataLabel.Text', 'ch.SeriesCollection(1).Points(1).DataLabel.Text', 'ch.SeriesCollection(1).Points(2).DataLabel.ShowValue',
    'ch.SeriesCollection(1).Points(2).DataLabel.ShowCategoryName', 'ch.SeriesCollection(1).Points(2).DataLabel.ShowSeriesName', 'ch.SeriesCollection(1).Points(2).DataLabel.ShowLegendKey',
    'ch.SeriesCollection(1).Points(2).DataLabel.ShowPercentage', 'ch.SeriesCollection(1).Points(2).DataLabel.Separator', 'ch.SeriesCollection(1).HasLeaderLines',
    'ch.SeriesCollection(1).DataLabels.ShowValue', 'ch.SeriesCollection(1).DataLabels.ShowCategoryName', 'ch.SeriesCollection(1).DataLabels.ShowSeriesName',
    'ch.SeriesCollection(1).DataLabels.ShowLegendKey', 'ch.SeriesCollection(1).DataLabels.ShowPercentage', 'ch.SeriesCollection(1).DataLabels.Separator', 'ch.SeriesCollection(1).DataLabels.Position',
    'ch.SeriesCollection(2).DataLabels.ShowValue',
]


def _probe() -> str:
    lines = ['Private Function Tried(ByVal which As Long, ByVal folder As String) As String',
             '    Dim wb As Workbook, ws As Worksheet, ch As Object, s As String, v As String', '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set ws = wb.Worksheets(1)',
             '    Set ch = ws.ChartObjects.Add(100, 10, 300, 200).Chart', '    ws.Range("A1:A4").Value = Application.Transpose(Array("k", "a", "b", "c"))',
             '    ws.Range("B1:B4").Value = Application.Transpose(Array("v", 1, 2, 3))', '    ws.Range("C1:C4").Value = Application.Transpose(Array("w", 4, 5, 6))',
             '    ch.SetSourceData ws.Range("A1:C4")', '    On Error GoTo Failed', '    Select Case which']
    for index, (owner, kind, before, argument) in enumerate(CASES):
        lines += [f'    Case {index}', f'        ch.ChartType = {TYPES[kind]}']
        lines += [f'        {BEFORE[before]}'] if before else []
        lines += [] if owner == NOBODY else [f'        {OWNERS[owner]}.ApplyDataLabels {argument}'.rstrip()]
    lines += ['    End Select', '    s = "ok "', 'Answered:', '    On Error Resume Next']
    for read in READS:
        lines += ['    Err.Clear', '    v = ""', f'    v = {read}', '    If Err.Number Then v = "E" & Err.Number', '    s = s & "," & v']
    lines += ['    On Error GoTo 0', '    Tried = s', '    wb.SaveAs folder & "\\c" & which & ".xlsx", xlOpenXMLWorkbook', '    wb.Close False', '    Exit Function',
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
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'reads': READS, 'source': source,
              'cases': [[f'{owner} {kind} {before}: {argument}', answer, part] for (owner, kind, before, argument), answer, part in zip(CASES, answers, parts, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for case, answer, part in record['cases']:
        found = re.findall(r'<c:dLbls>.*?</c:dLbls>', part, re.DOTALL)
        print(case, '\n  =', answer, '\n  +', ' || '.join(found)[:400])


if __name__ == '__main__':
    main()
