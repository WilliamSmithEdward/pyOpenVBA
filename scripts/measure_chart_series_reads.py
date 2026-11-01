"""What Series members read on a chart's first and second series, in each kind of chart: each member as its value,
TypeName for an object, or "E" and the error number.

Each case makes a workbook, adds a chart through ChartObjects.Add, fills categories in A2:A4 and two series in B and
C, gives them to the chart with SetSourceData, sets the case's chart type, then reads every member of NAMES on both
series.

    python scripts/measure_chart_series_reads.py

writes tests/fixtures/chart_series_reads.json, which tests/test_excel_chart_series_reads.py replays.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/chart_series_reads.json'
KINDS = ['xlColumnClustered', 'xlColumnStacked', 'xlBarClustered', 'xlLine', 'xlLineMarkers', 'xlPie', 'xlDoughnut',
         'xlArea', 'xlXYScatter']
NAMES = ['Parent', 'Application', 'Creator', 'AxisGroup', 'ChartType', 'Type', 'IsFiltered', 'HasErrorBars',
         'Has3DEffect', 'BarShape', 'PlotColorIndex', 'PictureType', 'PictureUnit2', 'InvertColorIndex', 'InvertColor',
         'ApplyPictToFront', 'ApplyPictToSides', 'ApplyPictToEnd', 'Shadow', 'ValueSortOrder',
         'QuartileCalculationInclusiveMedian', 'GeoMappingLevel', 'GeoProjectionType', 'RegionLabelOption',
         'ParentDataLabelOption', 'SeriesColorGradientStyle']


def _probe() -> str:
    lines = ['Private Function Safe(ByVal target As Object, ByVal name As String) As String',
             '    Dim v As Variant', '    On Error GoTo Failed',
             '    If IsObject(CallByName(target, name, VbGet)) Then',
             '        Safe = TypeName(CallByName(target, name, VbGet))', '    Else',
             '        v = CallByName(target, name, VbGet): Safe = TypeName(v) & ":" & CStr(v)', '    End If',
             '    Exit Function', 'Failed:', '    Safe = "E" & Err.Number', 'End Function', '',
             'Private Function Tried(ByVal kind As Long) As String',
             '    Dim wb As Workbook, ws As Worksheet, ch As Object, i As Long, j As Long, s As String, t As String',
             '    Dim names As Variant',
             *[f'    names = Array({", ".join(chr(34) + name + chr(34) for name in NAMES[:13])})' if part == 0 else
               f'    names = Split(Join(names, ",") & ",{",".join(NAMES[13:])}", ",")' for part in (0, 1)],
             '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set ws = wb.Worksheets(1)',
             '    Set ch = ws.ChartObjects.Add(100, 10, 300, 200).Chart',
             '    ws.Range("A1:A4").Value = Application.Transpose(Array("k", "a", "b", "c"))',
             '    ws.Range("B1:B4").Value = Application.Transpose(Array("v", 1, 2, 3))',
             '    ws.Range("C1:C4").Value = Application.Transpose(Array("w", 4, 5, 6))',
             '    ch.SetSourceData ws.Range("A1:C4")', '    ch.ChartType = kind',
             '    For i = 1 To 2',
             '        t = ""',
             '        For j = LBound(names) To UBound(names)',
             '            t = t & "," & Safe(ch.SeriesCollection(i), names(j))', '        Next',
             '        s = s & "|" & Mid(t, 2)', '    Next', '    Tried = Mid(s, 2)', '    wb.Close False', 'End Function', '',
             'Public Function Probe() As String', '    Dim s As String']
    lines += [f'    s = s & "@@@" & Tried({kind})' for kind in KINDS]
    lines += ['    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def main() -> None:
    source = _probe()
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(source, 'Probe', timeout=280.0)
        assert result.ok and not result.dialogs, result
        answers = str(result.value).split('@@@')
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source, 'names': NAMES,
              'cases': [list(pair) for pair in zip(KINDS, answers, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for kind, answer in record['cases']:
        for one in answer.split('|'):
            print(kind, dict(zip(NAMES, one.split(','), strict=True)))


if __name__ == '__main__':
    main()
