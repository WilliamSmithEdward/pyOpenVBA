"""A chart when the cells it reads are taken away: what it reads afterwards, what can be done with it, and the chart part
Excel saves.

Each case makes a workbook with categories in A2:A4 and two series in B and C on Data, and a chart of them on Data, of the
chart type and layout it names. It runs the case's statements, reads READS, each as its value or "E" and an error number,
and saves the workbook. The chart part is recorded as scripts/measure_chart_markup.py records it.

    python scripts/measure_chart_lost_cells.py

writes tests/fixtures/chart_lost_cells.json, which tests/test_excel_chart_lost_cells.py replays.
"""
from __future__ import annotations

import json
import tempfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

from measure_chart_markup import chart_part

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/chart_lost_cells.json'
TYPES = {'column': 'xlColumnClustered', 'bar': 'xlBarClustered', 'line': 'xlLine', 'scatter': 'xlXYScatter', 'pie': 'xlPie', 'area': 'xlArea'}
S1, S2 = 'ch.SeriesCollection(1)', 'ch.SeriesCollection(2)'
PAIR = 'ws.Range("E2:E4").Value = Application.Transpose(Array("p", "q", "r"))'
EDITS = [
    # One slot of one series, or of both.
    'd.Range("B1").Delete xlShiftToLeft', 'd.Range("B1").Delete xlShiftUp', 'd.Range("C1").Delete xlShiftUp', 'd.Range("B1:C1").Delete xlShiftUp', 'd.Range("A2:A4").Delete xlShiftUp', 'd.Range("A2:A4").Delete xlShiftToLeft',
    'd.Range("B2:B4").Delete xlShiftUp', 'd.Range("B2:C4").Delete xlShiftUp', 'd.Range("A2:C4").Delete xlShiftUp', 'd.Range("A1:B4").Delete xlShiftToLeft', 'd.Range("A1:C1").Delete xlShiftUp',
    'd.Rows(1).Delete', 'd.Rows("2:4").Delete', 'd.Rows("1:4").Delete', 'd.Columns(1).Delete', 'd.Columns(2).Delete', 'd.Columns(3).Delete', 'd.Columns("B:C").Delete', 'd.Columns("A:C").Delete', 'd.Columns("A:B").Delete', 'd.Delete',
    # Each part of a reference: part of it going takes nothing away.
    'd.Rows(2).Delete', 'd.Rows("2:3").Delete', 'd.Rows("3:9").Delete',
    # Series with their own categories, names and values.
    f'{S2}.XValues = d.Range("E2:E4"): d.Columns(1).Delete', f'{S2}.XValues = d.Range("E2:E4"): d.Columns(5).Delete', f'{S2}.XValues = d.Range("E2:E4"): d.Rows(1).Delete', f'{S1}.Name = "z": d.Rows(1).Delete', f'{S1}.Name = "z": d.Columns(2).Delete', f'{S1}.Name = "z": {S2}.Name = "y": d.Rows("2:4").Delete',
    f'{S1}.Values = Array(7, 8, 9): d.Columns(2).Delete', f'{S1}.XValues = Array("x", "y", "z"): d.Columns(1).Delete', f'{S1}.Name = "=Data!$E$1": d.Columns(5).Delete', f'{S1}.Name = "=Data!$E$1": d.Rows(1).Delete',
    # What comes after.
    f'd.Columns(2).Delete: {S1}.Values = d.Range("B2:B4")', f'd.Columns(2).Delete: {S1}.Name = "z"', f'd.Columns(2).Delete: {S1}.XValues = d.Range("B2:B4")', f'd.Columns(2).Delete: {S1}.Delete', 'd.Columns(2).Delete: d.Rows(1).Insert', 'd.Columns(2).Delete: d.Columns(1).Insert', 'd.Columns(2).Delete: ch.SetSourceData d.Range("A1:B4")', 'd.Columns(2).Delete: d.Columns(2).Delete',
    'd.Columns(2).Delete: ch.ChartType = xlLine', f'd.Columns(2).Delete: {S1}.HasDataLabels = True', f'd.Columns(2).Delete: {S1}.Points(1).Format.Fill.ForeColor.RGB = 255', f'd.Columns(2).Delete: {S1}.Format.Fill.ForeColor.RGB = 255', f'd.Columns(2).Delete: {S1}.Trendlines.Add',
    f'd.Rows(1).Delete: {S1}.Name = "z"', f'd.Rows(1).Delete: {S1}.Name = "=Data!$A$1"', f'd.Rows(1).Delete: {S1}.Values = d.Range("B1:B3")', 'd.Rows(1).Delete: d.Rows(1).Insert', 'd.Rows(1).Delete: d.Rows(1).Delete',
    f'd.Columns(1).Delete: {S1}.XValues = d.Range("A2:A4")', f'd.Columns(1).Delete: {S1}.XValues = Array("x", "y", "z")', 'd.Columns(1).Delete: d.Columns(1).Insert', 'd.Columns(1).Delete: d.Columns(1).Delete', f'd.Rows("2:4").Delete: {S1}.Values = d.Range("B1:B1")', 'd.Rows("1:4").Delete: ch.SetSourceData d.Range("A1:B2")',
    # What the series holds.
    f'{S1}.Points(2).Format.Fill.ForeColor.RGB = 255: d.Columns(2).Delete', f'{S1}.Points(1).Format.Fill.ForeColor.RGB = 255: d.Columns(2).Delete', f'{S1}.Points(2).Format.Fill.ForeColor.RGB = 255: d.Rows("2:4").Delete', f'{S1}.Points(2).Format.Fill.ForeColor.RGB = 255: d.Columns(1).Delete', f'{S1}.HasDataLabels = True: d.Columns(2).Delete', f'{S1}.Points(2).HasDataLabel = True: d.Columns(2).Delete', f'{S1}.HasDataLabels = True: d.Rows(1).Delete',
    f'{S1}.Trendlines.Add: d.Columns(2).Delete', f'{S1}.Format.Fill.ForeColor.RGB = 255: d.Columns(2).Delete', f'{S1}.ErrorBar xlY, xlBoth, xlCustom, d.Range("E2:E4"), d.Range("E2:E4"): d.Columns(5).Delete', f'{S1}.ErrorBar xlY, xlBoth, xlCustom, d.Range("E2:E4"), d.Range("E2:E4"): d.Columns(2).Delete',
    # Titles that read cells.
    'ch.HasTitle = True: ch.ChartTitle.Formula = "=Data!$A$1": d.Rows(1).Delete', 'ch.HasTitle = True: ch.ChartTitle.Formula = "=Data!$E$1": d.Columns(5).Delete', 'ch.HasTitle = True: ch.ChartTitle.Formula = "=Data!$E$1": d.Columns(5).Delete: ch.ChartTitle.Text = "t"', 'ch.HasTitle = True: ch.ChartTitle.Formula = "=Data!$E$1": d.Columns(5).Delete: ch.ChartTitle.Formula = "=Data!$A$1"',
    'ch.Axes(xlValue).HasTitle = True: ch.Axes(xlValue).AxisTitle.Formula = "=Data!$E$1": d.Columns(5).Delete', 'ch.Axes(xlCategory).HasTitle = True: ch.Axes(xlCategory).AxisTitle.Formula = "=Data!$A$1": d.Rows(1).Delete',
]
CASES = [('column', 'columns', edit) for edit in EDITS]
LOSSES = ('d.Rows(1).Delete', 'd.Rows("2:4").Delete', 'd.Rows("1:4").Delete', 'd.Columns(1).Delete', 'd.Columns(2).Delete', 'd.Columns(3).Delete', 'd.Columns("B:C").Delete', 'd.Range("B1").Delete xlShiftUp', 'd.Range("B2:B4").Delete xlShiftUp')
CASES += [(kind, 'columns', edit) for kind in ('bar', 'line', 'scatter', 'pie', 'area') for edit in LOSSES]
CASES += [('column', 'rows', edit) for edit in ('d.Rows(1).Delete', 'd.Rows(2).Delete', 'd.Rows(3).Delete', 'd.Rows("2:3").Delete', 'd.Rows("1:3").Delete', 'd.Columns(1).Delete', 'd.Columns(2).Delete', 'd.Columns("B:C").Delete', 'd.Range("A2").Delete xlShiftUp')]
READS = ['ch.SeriesCollection.Count', f'{S1}.Formula', f'{S2}.Formula', f'{S1}.Name', f'{S2}.Name', f'{S1}.Points.Count', f'{S2}.Points.Count', f'UBound({S1}.Values)', f'CStr({S1}.Values()(1))', f'UBound({S1}.XValues)', f'CStr({S1}.XValues()(1))', f'{S1}.IsFiltered', 'ch.FullSeriesCollection.Count',
         'ch.PlotBy', 'ch.Legend.LegendEntries.Count', 'ch.HasTitle', 'ch.ChartTitle.Text', 'ch.ChartTitle.Formula', f'{S1}.Points(1).DataLabel.Text', f'{S1}.HasDataLabels']


def _probe() -> str:
    lines = ['Private Function Tried(ByVal which As Long, ByVal folder As String) As String',
             '    Dim wb As Workbook, d As Worksheet, o As Worksheet, ch As Object, s As String, v As String', '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set d = wb.Worksheets(1): d.Name = "Data"', '    Set o = wb.Worksheets.Add(After:=d): o.Name = "Other": d.Activate',
             '    d.Range("A1:A4").Value = Application.Transpose(Array("k", "a", "b", "c"))', '    d.Range("B1:B4").Value = Application.Transpose(Array("v", 1, 2, 3))', '    d.Range("C1:C4").Value = Application.Transpose(Array("w", 4, 5, 6))', '    d.Range("E1:E4").Value = Application.Transpose(Array("e", 7, 8, 9))',
             '    Set ch = o.ChartObjects.Add(300, 100, 300, 200).Chart', '    On Error GoTo Failed', '    Select Case which']
    for index, (kind, layout, edit) in enumerate(CASES):
        lines += [f'    Case {index}', f'        ch.SetSourceData d.Range("A1:C4"), xl{layout.capitalize()}', f'        ch.ChartType = {TYPES[kind]}', f'        {edit}']
    lines += ['    End Select', '    s = "ok "', 'Answered:', '    On Error Resume Next']
    for read in READS:
        lines += ['    Err.Clear', '    v = ""', f'    v = {read}', '    If Err.Number Then v = "E" & Err.Number', '    s = s & "|" & v']
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
              'cases': [[f'{kind} {layout}: {edit}', answer, part] for (kind, layout, edit), answer, part in zip(CASES, answers, parts, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for case, answer, _ in record['cases']:
        print(case.replace(S1, 'S1').replace(S2, 'S2'), '=>', answer)


if __name__ == '__main__':
    main()
