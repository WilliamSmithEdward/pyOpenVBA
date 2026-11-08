"""What Excel's save does to a loaded chart's part for one structural edit at a time, and what it answers after: series added and
deleted, source data, chart type and plot order. On three books Excel saved: the plain and dressed books of
scripts/measure_loaded_chart_edits.py and the labelled book of scripts/measure_loaded_chart_labels.py. Each case opens its book in a new
Excel session, as the model opens a file. Needs Excel:

    python scripts/measure_loaded_chart_structure_edits.py

writes tests/fixtures/loaded_chart_structure_edits.json, which tests/test_excel_loaded_chart_structure_edits.py replays.
"""

from __future__ import annotations

import datetime
import json
import shutil
import tempfile
import zipfile
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig
from pyvbaharness.results import HarnessError

FIXTURES = Path(__file__).resolve().parents[1] / 'tests/fixtures'
BOOKS = ('loaded_chart_plain.xlsx', 'loaded_chart_dressed.xlsx', 'loaded_chart_labelled.xlsx')
SC, S1, S2 = 'c.SeriesCollection', 'c.SeriesCollection(1)', 'c.SeriesCollection(2)'
SHEET = 'c.Parent.Parent'
#: Each: what it is, and the statements, with c the chart. A read follows in r, for its answer and type.
EDITS = (
    ('nothing', ''),
    ('new series', f'{SC}.NewSeries: r = {SC}.Count'),
    ('new series twice', f'{SC}.NewSeries: {SC}.NewSeries: r = {SC}.Count'),
    ('new series filled', f'With {SC}.NewSeries: .Name = "=One!$A$1": .Values = "=One!$B$2:$C$2": End With: r = {SC}(3).Formula'),
    ('delete last', f'{S2}.Delete: r = {SC}.Count'),
    ('delete first', f'{S1}.Delete: r = {S1}.Formula'),
    ('delete first then new', f'{S1}.Delete: {SC}.NewSeries: r = {SC}.Count'),
    ('new then delete it', f'{SC}.NewSeries: {SC}(3).Delete: r = {SC}.Count'),
    ('source same', f'c.SetSourceData {SHEET}.Range("A1:C3"): r = {SC}.Count'),
    ('source narrower', f'c.SetSourceData {SHEET}.Range("A1:B3"): r = {S1}.Formula'),
    ('source by columns', f'c.SetSourceData {SHEET}.Range("A1:C3"), xlColumns: r = {S1}.Formula'),
    ('plot by columns', f'c.PlotBy = xlColumns: r = {S1}.Formula'),
    ('series values', f'{S1}.Values = "=One!$B$3:$C$3": r = {S1}.Formula'),
    ('series categories', f'{S1}.XValues = "=One!$A$2:$A$3": r = {S1}.Formula'),
    ('series formula', f'{S1}.Formula = "=SERIES(One!$A$3,One!$B$1:$C$1,One!$B$3:$C$3,1)": r = {S1}.Formula'),
    ('type line', f'c.ChartType = xlLine: r = c.ChartType'),
    ('type line markers', f'c.ChartType = xlLineMarkers: r = c.ChartType'),
    ('type bar', f'c.ChartType = xlBarClustered: r = c.ChartType'),
    ('type stacked', f'c.ChartType = xlColumnStacked: r = c.ChartType'),
    ('type area', f'c.ChartType = xlArea: r = c.ChartType'),
    ('type pie', f'c.ChartType = xlPie: r = c.ChartType'),
    ('type line and back', f'c.ChartType = xlLine: c.ChartType = xlColumnClustered: r = c.ChartType'),
    ('type same', f'c.ChartType = xlColumnClustered: r = c.ChartType'),
    ('series type line', f'{S2}.ChartType = xlLine: r = {S2}.ChartType'),
)


def source() -> str:
    lines = ['Private Sub Edit(ByVal which As Long, c As Object, r As Variant)', '    Select Case which']
    lines += [f'    Case {index}: {statements}' for index, (_, statements) in enumerate(EDITS) if statements]
    lines += ['    End Select', 'End Sub', r'''
Public Function Act(ByVal path As String, ByVal out As String, ByVal which As Long) As String
    Dim wb As Workbook, c As Object, r As Variant, s As String
    Set wb = Workbooks.Open(path)
    Set c = wb.Worksheets(1).ChartObjects(1).Chart
    On Error Resume Next
    Edit which, c, r
    s = Err.Number & " " & TypeName(r) & " " & r
    On Error GoTo 0
    Application.DisplayAlerts = False
    wb.SaveAs out, xlOpenXMLWorkbook
    Application.DisplayAlerts = True
    wb.Close False
    Act = s
End Function
''']
    return '\n'.join(lines)


def chart_parts(path: Path) -> dict[str, str]:
    with zipfile.ZipFile(path) as package:
        return {name: package.read(name).decode('utf-8') for name in package.namelist() if name.startswith('xl/charts/') and name.endswith('.xml')}


def main() -> None:
    folder = Path(tempfile.mkdtemp())
    code = source()
    cases: list[list[object]] = []
    originals = {name: chart_parts(FIXTURES / name) for name in BOOKS}
    for name in BOOKS:
        for which, (label, _) in enumerate(EDITS):
            answer, saved = '', {}
            for _turn in range(3):
                copy, out = folder / 'in.xlsx', folder / f'{name}{which}.xlsx'
                shutil.copy(FIXTURES / name, copy)
                try:
                    with ExcelSession(HarnessConfig(lock_wait_s=600)) as excel:
                        done = excel.run_vba(code, 'Act', (str(copy), str(out), which), timeout=120.0, line_numbers=False)
                except HarnessError as failed:
                    # Excel's worker can be slow to start; the case runs again in a new session.
                    answer = f'harness: {failed}'
                    continue
                answer = done.value if done.outcome == 'passed' else f'{done.outcome}: {done.message}'
                if done.outcome == 'passed':
                    saved = chart_parts(out)
                    break
            cases.append([name, which, answer, saved])
            print(name, label, '=>', answer, flush=True)
    record = {'measured_at': datetime.datetime.now(datetime.timezone.utc).isoformat(), 'source': code, 'edits': EDITS, 'originals': originals,
              'cases': cases}
    (FIXTURES / 'loaded_chart_structure_edits.json').write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8', newline='\n')


if __name__ == '__main__':
    main()
