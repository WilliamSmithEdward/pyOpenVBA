"""What Excel's save keeps of a loaded chart's data labels when its data changes: SetSourceData, Series.Values and
Series.Formula, on the labelled book and on the plain book with labels set in the session. Needs Excel:

    python scripts/measure_loaded_chart_label_data.py

writes tests/fixtures/loaded_chart_label_data.json.
"""

from __future__ import annotations

import datetime
import json
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig
from pyvbaharness.results import HarnessError

FIXTURES = Path(__file__).resolve().parents[1] / 'tests/fixtures'
BOOKS = ('loaded_chart_labelled.xlsx', 'loaded_chart_plain.xlsx')
S1, S2, SHEET = 'c.SeriesCollection(1)', 'c.SeriesCollection(2)', 'c.Parent.Parent'
#: Labels the plain book's chart takes first, so that both books have them.
LABELS = f'{S1}.HasDataLabels = True: {S1}.DataLabels.ShowCategoryName = True: {S2}.Points(2).HasDataLabel = True: '
#: Each: what it is, and the statements, with c the chart. A read follows in r, for its answer and type.
EDITS = (
    ('nothing', ''),
    ('source same', f'c.SetSourceData {SHEET}.Range("A1:C3"): r = c.SeriesCollection.Count'),
    ('source first row', f'c.SetSourceData {SHEET}.Range("A1:C2"): r = {S1}.Formula'),
    ('source same by rows', f'c.SetSourceData {SHEET}.Range("A1:C3"), xlRows: r = c.SeriesCollection.Count'),
    ('source narrower', f'c.SetSourceData {SHEET}.Range("A1:B3"): r = {S1}.Formula'),
    ('source by columns', f'c.SetSourceData {SHEET}.Range("A1:C3"), xlColumns: r = {S1}.Formula'),
    ('source columns and back', f'c.SetSourceData {SHEET}.Range("A1:C3"), xlColumns: c.SetSourceData {SHEET}.Range("A1:C3"), xlRows: '
                                f'r = {S1}.Formula'),
    ('plot by columns', f'c.PlotBy = xlColumns: r = {S1}.Formula'),
    ('values same', f'{S1}.Values = "=One!$B$2:$C$2": r = {S1}.Formula'),
    ('values other', f'{S1}.Values = "=One!$B$3:$C$3": r = {S1}.Formula'),
    ('values second', f'{S2}.Values = "=One!$B$2:$C$2": r = {S2}.Formula'),
    ('values array', f'{S1}.Values = Array(1, 2): r = {S1}.Formula'),
    ('formula same', f'{S1}.Formula = {S1}.Formula: r = {S1}.Formula'),
    ('formula name', f'{S1}.Formula = "=SERIES(""N"",One!$B$1:$C$1,One!$B$2:$C$2,1)": r = {S1}.Formula'),
    ('formula values', f'{S1}.Formula = "=SERIES(One!$A$2,One!$B$1:$C$1,One!$B$3:$C$3,1)": r = {S1}.Formula'),
)


def source() -> str:
    lines = ['Private Sub Edit(ByVal which As Long, c As Object, r As Variant)', '    Select Case which']
    lines += [f'    Case {index}: {statements}' for index, (_, statements) in enumerate(EDITS) if statements]
    lines += ['    End Select', 'End Sub', r'''
Public Function Act(ByVal path As String, ByVal out As String, ByVal which As Long, ByVal first As String) As String
    Dim wb As Workbook, c As Object, r As Variant, s As String
    Set wb = Workbooks.Open(path)
    Set c = wb.Worksheets(1).ChartObjects(1).Chart
    If first <> "" Then Labelled c
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

Public Sub Labelled(c As Object)
    ''' + LABELS.rstrip(': ') + r'''
End Sub
''']
    return '\n'.join(lines)


def chart_parts(path: Path) -> dict[str, str]:
    with zipfile.ZipFile(path) as package:
        return {name: package.read(name).decode('utf-8') for name in package.namelist() if name == 'xl/charts/chart1.xml'}


def main() -> None:
    folder = Path(tempfile.mkdtemp())
    code = source()
    cases: list[list[object]] = []
    originals = {name: chart_parts(FIXTURES / name) for name in BOOKS}
    for name in BOOKS:
        first = '' if name == BOOKS[0] else 'labels'
        for which, (label, _) in enumerate(EDITS):
            answer, saved = '', {}
            for _turn in range(2):
                copy, out = folder / 'in.xlsx', folder / f'{name}{which}.xlsx'
                shutil.copy(FIXTURES / name, copy)
                try:
                    with ExcelSession(HarnessConfig(lock_wait_s=600)) as excel:
                        done = excel.run_vba(code, 'Act', (str(copy), str(out), which, first), timeout=120.0, line_numbers=False)
                except HarnessError as failed:
                    answer = f'harness: {failed}'
                    continue
                answer = done.value if done.outcome == 'passed' else f'{done.outcome}: {done.message}'
                if done.outcome == 'passed':
                    saved = chart_parts(out)
                    break
            cases.append([name, which, answer, saved])
            print(name, label, '=>', answer, flush=True)
    record = {'measured_at': datetime.datetime.now(datetime.timezone.utc).isoformat(), 'source': code, 'edits': EDITS,
              'labels': LABELS, 'originals': originals, 'cases': cases}
    (FIXTURES / 'loaded_chart_label_data.json').write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')


if __name__ == '__main__':
    sys.exit(main())
