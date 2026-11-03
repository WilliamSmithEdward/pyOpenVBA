"""A chart's series when the cells it reads are hidden: rows, columns and filtered rows, with PlotVisibleOnly on and off.
What the series reads afterwards, and the chart part Excel saves.

Each case makes a workbook with categories in A2:A5 and two series in B and C, and a chart of them, of the chart type TYPES
names and laid out by columns or by rows. It runs the case's statements, reads READS, each as its value or "E" and an error
number, and saves the workbook. The chart part is recorded as scripts/measure_chart_markup.py records it.

    python scripts/measure_chart_hidden_cells.py

writes tests/fixtures/chart_hidden_cells.json, which tests/test_excel_chart_hidden_cells.py replays.
"""
from __future__ import annotations

import json
import tempfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

from measure_chart_markup import chart_part

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/chart_hidden_cells.json'
TYPES = {'column': 'xlColumnClustered', 'line': 'xlLine', 'scatter': 'xlXYScatter', 'pie': 'xlPie'}
S1, S2 = 'ch.SeriesCollection(1)', 'ch.SeriesCollection(2)'
EDITS = [
    '', 'ws.Rows(1).Hidden = True', 'ws.Rows(2).Hidden = True', 'ws.Rows(3).Hidden = True', 'ws.Rows(5).Hidden = True', 'ws.Rows("3:4").Hidden = True', 'ws.Rows("2:5").Hidden = True', 'ws.Rows(9).Hidden = True',
    'ws.Columns(1).Hidden = True', 'ws.Columns(2).Hidden = True', 'ws.Columns(3).Hidden = True', 'ws.Columns("B:C").Hidden = True', 'ws.Columns(1).Hidden = True: ws.Columns(2).Hidden = True',
    'ws.Rows(3).Hidden = True: ws.Rows(3).Hidden = False', 'ws.Rows(3).RowHeight = 0', 'ws.Columns(2).ColumnWidth = 0', 'ws.Range("A1:C5").AutoFilter 2, ">2"', 'ws.Range("A1:C5").AutoFilter 2, ">2": ws.ShowAllData', 'ws.Range("A1:C5").AutoFilter 2, ">9"',
    'ch.PlotVisibleOnly = False: ws.Rows(3).Hidden = True', 'ws.Rows(3).Hidden = True: ch.PlotVisibleOnly = False', 'ws.Rows(3).Hidden = True: ch.PlotVisibleOnly = False: ch.PlotVisibleOnly = True', 'ch.PlotVisibleOnly = False: ws.Columns(2).Hidden = True',
    f'ws.Rows(3).Hidden = True: {S1}.Points(2).Format.Fill.ForeColor.RGB = 255', f'{S1}.Points(2).Format.Fill.ForeColor.RGB = 255: ws.Rows(3).Hidden = True', f'{S1}.Points(3).Format.Fill.ForeColor.RGB = 255: ws.Rows(3).Hidden = True', f'{S1}.Points(3).Format.Fill.ForeColor.RGB = 255: ws.Rows(3).Hidden = True: ws.Rows(3).Hidden = False',
    f'ws.Rows(3).Hidden = True: {S1}.Values = ws.Range("B2:B5")', f'{S1}.Values = Array(7, 8, 9, 10): ws.Rows(3).Hidden = True', f'{S1}.XValues = Array("x", "y", "z", "t"): ws.Rows(3).Hidden = True', f'{S1}.Name = "z": ws.Rows(1).Hidden = True',
    f'ws.Rows(3).Hidden = True: {S1}.HasDataLabels = True', f'ws.Rows(3).Hidden = True: {S1}.Points(2).HasDataLabel = True', 'ws.Rows(3).Hidden = True: ws.Rows(3).Delete', 'ws.Rows(3).Hidden = True: ws.Rows(2).Insert', 'ws.Rows(3).Hidden = True: ws.Range("B3").Value = 9', 'ws2.Rows(3).Hidden = True',
    # What comes after a series or its categories are all hidden.
    'ws.Columns(2).Hidden = True: ws.Columns(2).Hidden = False', 'ws.Columns(1).Hidden = True: ws.Columns(1).Hidden = False', 'ws.Rows("2:5").Hidden = True: ws.Rows("2:5").Hidden = False', 'ws.Columns(2).Hidden = True: ch.PlotVisibleOnly = False', 'ws.Columns(1).Hidden = True: ch.PlotVisibleOnly = False',
    'ws.Columns(2).Hidden = True: ws.Rows(1).Insert', 'ws.Columns(2).Hidden = True: ws.Columns(2).Delete', 'ws.Columns(2).Hidden = True: ws.Columns(1).Delete', f'ws.Columns(2).Hidden = True: {S1}.Name = "z"', f'ws.Columns(2).Hidden = True: {S1}.Name = "z": ws.Columns(2).Hidden = False',
    'ws.Columns(2).Hidden = True: ch.SeriesCollection.NewSeries', 'ws.Columns(2).Hidden = True: ch.SetSourceData ws.Range("A1:B5")', 'ws.Columns(2).Hidden = True: ch.ChartType = xlLine', f'{S1}.Format.Fill.ForeColor.RGB = 255: ws.Columns(2).Hidden = True', f'{S1}.Format.Fill.ForeColor.RGB = 255: ws.Columns(2).Hidden = True: ws.Columns(2).Hidden = False',
    f'{S1}.Points(2).Format.Fill.ForeColor.RGB = 255: ws.Rows(3).Hidden = True: ws.Rows(3).Hidden = False', f'{S1}.Points(2).HasDataLabel = True: ws.Rows(3).Hidden = True', f'{S1}.Points(3).HasDataLabel = True: ws.Rows(3).Hidden = True', f'{S1}.Points(2).Format.Fill.ForeColor.RGB = 255: ws.Columns(2).Hidden = True: ws.Columns(2).Hidden = False',
]
CASES = [('column', 'columns', edit) for edit in EDITS]
CASES += [(kind, 'columns', edit) for kind in ('line', 'scatter', 'pie') for edit in ('ws.Rows(3).Hidden = True', 'ws.Columns(1).Hidden = True', 'ws.Columns(2).Hidden = True', 'ws.Rows(1).Hidden = True')]
CASES += [('column', 'rows', edit) for edit in ('', 'ws.Rows(1).Hidden = True', 'ws.Rows(2).Hidden = True', 'ws.Columns(1).Hidden = True', 'ws.Columns(2).Hidden = True', 'ws.Columns(3).Hidden = True')]
READS = ['ch.SeriesCollection.Count', f'{S1}.Name', f'{S1}.Points.Count', f'{S2}.Points.Count', f'{S1}.Formula', f'UBound({S1}.Values)', f'{S1}.Values(1)', f'UBound({S1}.XValues)', f'CStr({S1}.XValues()(1))', f'CStr({S1}.Values()(1))', f'CStr({S1}.Values()(2))',
         f'{S1}.Points(2).DataLabel.Text', f'{S1}.Points(2).HasDataLabel', 'ch.PlotVisibleOnly', 'ch.PlotBy', 'ch.Legend.LegendEntries.Count', 'ch.FullSeriesCollection.Count', f'{S1}.IsFiltered']


def _probe() -> str:
    lines = ['Private Function Tried(ByVal which As Long, ByVal folder As String) As String',
             '    Dim wb As Workbook, ws As Worksheet, ws2 As Worksheet, ch As Object, s As String, v As String', '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set ws = wb.Worksheets(1)', '    Set ws2 = wb.Worksheets.Add(After:=ws): ws.Activate',
             '    ws.Range("A1:A5").Value = Application.Transpose(Array("k", "a", "b", "c", "d"))', '    ws.Range("B1:B5").Value = Application.Transpose(Array("v", 1, 2, 3, 4))', '    ws.Range("C1:C5").Value = Application.Transpose(Array("w", 5, 6, 7, 8))',
             '    Set ch = ws.ChartObjects.Add(300, 100, 300, 200).Chart', '    On Error GoTo Failed', '    Select Case which']
    for index, (kind, layout, edit) in enumerate(CASES):
        lines += [f'    Case {index}', f'        ch.SetSourceData ws.Range("A1:C5"), xl{layout.capitalize()}', f'        ch.ChartType = {TYPES[kind]}']
        lines += [f'        {edit}'] if edit else []
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
        print(case.replace(S1, 'S1'), '=>', answer)


if __name__ == '__main__':
    main()
