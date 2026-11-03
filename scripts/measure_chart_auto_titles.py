"""When a chart gets Excel's automatic title: by how many series it draws, how they came, and where its data is. What it reads,
and the chart part Excel saves.

Each case makes a workbook with categories in A2:A4 and series in B, C and D on Data, a sheet Other, and an empty chart on the
sheet it names, and runs the case's statements. It does so twice, since reading ChartTitle may itself change the chart: once
saving at once, and once reading READS first, each as its value or "E" and an error number. Each chart part is recorded as
scripts/measure_chart_markup.py records it.

    python scripts/measure_chart_auto_titles.py

writes tests/fixtures/chart_auto_titles.json, which tests/test_excel_chart_auto_titles.py replays.
"""
from __future__ import annotations

import json
import tempfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

from measure_chart_markup import chart_part

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/chart_auto_titles.json'
S = 'ch.SeriesCollection'
STATEMENTS = [
    '', 'ch.SetSourceData d.Range("A1:B4")', 'ch.SetSourceData d.Range("A1:C4")', 'ch.SetSourceData d.Range("B1:B4")', 'ch.SetSourceData d.Range("B2:B4")', 'ch.SetSourceData d.Range("A1:B4"), xlRows', 'ch.SetSourceData d.Range("A1:D2"), xlRows', 'ch.SetSourceData d.Range("A1:D2"), xlColumns',
    'ch.SetSourceData d.Range("A1:C4"): ch.SetSourceData d.Range("A1:B4")', 'ch.SetSourceData d.Range("A1:B4"): ch.SetSourceData d.Range("A1:C4")', 'ch.SetSourceData d.Range("A1:B4"): ch.HasTitle = False: ch.SetSourceData d.Range("A1:C4")', 'ch.SetSourceData d.Range("A1:B4"): ch.HasTitle = False: ch.SetSourceData d.Range("A1:B4")', 'ch.SetSourceData d.Range("A1:C4"): ch.HasTitle = True: ch.SetSourceData d.Range("A1:B4")', 'ch.SetSourceData d.Range("A1:B4"): ch.ChartTitle.Text = "t": ch.SetSourceData d.Range("A1:C4")',
    f'{S}.NewSeries', f'{S}.NewSeries: {S}(1).Values = d.Range("B2:B4")', f'{S}.NewSeries: {S}.NewSeries', f'{S}.Add d.Range("B1:B4")', f'{S}.Add d.Range("B1:C4")', f'{S}.Add d.Range("B1:B4"): {S}.Add d.Range("C1:C4")', f'{S}.Extend d.Range("B1:B4")',
    f'ch.SetSourceData d.Range("A1:C4"): {S}(2).Delete', f'ch.SetSourceData d.Range("A1:B4"): {S}(1).Delete', f'ch.SetSourceData d.Range("A1:B4"): {S}.NewSeries', f'ch.SetSourceData d.Range("A1:B4"): {S}.Add d.Range("C1:C4")', f'ch.SetSourceData d.Range("A1:C4"): {S}(1).IsFiltered = True', f'ch.SetSourceData d.Range("A1:C4"): {S}(2).Delete: {S}(1).Delete', f'ch.SetSourceData d.Range("A1:B4"): {S}(1).Name = "z"',
    'ch.ChartType = xlPie: ch.SetSourceData d.Range("A1:B4")', 'ch.ChartType = xlPie: ch.SetSourceData d.Range("A1:C4")', 'ch.SetSourceData d.Range("A1:B4"): ch.ChartType = xlPie', 'ch.SetSourceData d.Range("A1:C4"): ch.ChartType = xlPie', 'ch.SetSourceData d.Range("A1:B4"): ch.ChartType = xlLine', 'ch.SetSourceData d.Range("A1:B4"): ch.ChartType = xlXYScatter',
    'd.Columns(3).Hidden = True: ch.SetSourceData d.Range("A1:C4")', 'ch.SetSourceData d.Range("A1:C4"): d.Columns(3).Hidden = True', 'ch.SetSourceData d.Range("A1:B4"): d.Range("B1").Value = "changed"', 'ch.SetSourceData d.Range("A1:C4"): d.Columns(3).Delete',
]
CASES = [(place, statement) for place in ('Data', 'Other') for statement in STATEMENTS]
READS = ['ch.HasTitle', 'ch.ChartTitle.Text', f'{S}.Count', 'ch.HasLegend', 'ch.ChartTitle.Formula']


def _probe() -> str:
    lines = ['Private Function Tried(ByVal which As Long, ByVal folder As String, ByVal reading As Boolean) As String',
             '    Dim wb As Workbook, d As Object, o As Object, ch As Object, s As String, v As String', '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set d = wb.Worksheets(1): d.Name = "Data"', '    Set o = wb.Worksheets.Add(After:=d): o.Name = "Other": d.Activate',
             '    d.Range("A1:A4").Value = Application.Transpose(Array("k", "a", "b", "c"))', '    d.Range("B1:B4").Value = Application.Transpose(Array("v", 1, 2, 3))', '    d.Range("C1:C4").Value = Application.Transpose(Array("w", 4, 5, 6))', '    d.Range("D1:D4").Value = Application.Transpose(Array("x", 7, 8, 9))',
             '    d.Range("H1").Select', '    On Error GoTo Failed', '    Select Case which']
    for index, (place, statement) in enumerate(CASES):
        lines += [f'    Case {index}', f'        Set ch = wb.Worksheets("{place}").ChartObjects.Add(300, 100, 300, 200).Chart']
        lines += [f'        {statement}'] if statement else []
    lines += ['    End Select', '    s = "ok "', 'Answered:', '    On Error Resume Next', '    If reading Then']
    for read in READS:
        lines += ['    Err.Clear', '    v = ""', f'    v = {read}', '    If Err.Number Then v = "E" & Err.Number', '    s = s & "|" & v']
    lines += ['    End If', '    On Error GoTo 0', '    Tried = s', '    wb.SaveAs folder & "\\" & IIf(reading, "r", "c") & which & ".xlsx", xlOpenXMLWorkbook', '    wb.Close False', '    Exit Function',
              'Failed:', '    s = "Err " & Err.Number & " "', '    Resume Answered', 'End Function', '',
              'Public Function Probe(ByVal folder As String, ByVal first As Long, ByVal last As Long) As String', '    Dim i As Long, s As String',
              '    Application.DisplayAlerts = False', '    For i = first To last', '        s = s & "@@@" & Tried(i, folder, True)', '        Tried i, folder, False', '    Next', '    Application.DisplayAlerts = True',
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
        read = [chart_part(Path(folder) / f'r{index}.xlsx') for index in range(len(CASES))]
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'reads': READS, 'source': source,
              'cases': [[f'{place}: {statement}', answer, part, after] for (place, statement), answer, part, after in zip(CASES, answers, parts, read, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for case, answer, part, after in record['cases']:
        print(case, '=>', answer, 'unread', '<c:title>' in part, 'read', '<c:title>' in after)


if __name__ == '__main__':
    main()
