"""Which series take the categories a loaded chart's Series.Formula changes: on the plain book, and on one whose second
series has categories of its own. Needs Excel:

    python scripts/measure_loaded_chart_formula_categories.py

writes tests/fixtures/loaded_chart_formula_categories.json and tests/fixtures/loaded_chart_own_categories.xlsx.
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

sys.path.insert(0, str(Path(__file__).resolve().parent))
from measure_loaded_chart_edits import without_folder  # noqa: E402

FIXTURES = Path(__file__).resolve().parents[1] / 'tests/fixtures'
OWN = 'loaded_chart_own_categories.xlsx'
BOOKS = ('loaded_chart_plain.xlsx', OWN)
S1, S2 = 'c.SeriesCollection(1)', 'c.SeriesCollection(2)'
#: Each: what it is, and the statements, with c the chart. A read follows in r, for its answer and type.
EDITS = (
    ('nothing', ''),
    ('first new categories', f'{S1}.Formula = "=SERIES(One!$A$2,One!$A$2:$A$3,One!$B$2:$C$2,1)": r = {S1}.Formula & "|" & {S2}.Formula'),
    ('second new categories', f'{S2}.Formula = "=SERIES(One!$A$3,One!$A$2:$A$3,One!$B$3:$C$3,2)": r = {S1}.Formula & "|" & {S2}.Formula'),
    ('first other values', f'{S1}.Formula = "=SERIES(One!$A$2,One!$B$1:$C$1,One!$B$3:$C$3,1)": r = {S1}.Formula & "|" & {S2}.Formula'),
    ('first xvalues', f'{S1}.XValues = "=One!$A$2:$A$3": r = {S1}.Formula & "|" & {S2}.Formula'),
)


def source() -> str:
    lines = ['Private Sub Edit(ByVal which As Long, c As Object, r As Variant)', '    Select Case which']
    lines += [f'    Case {index}: {statements}' for index, (_, statements) in enumerate(EDITS) if statements]
    lines += ['    End Select', 'End Sub', r'''
Public Function Make(ByVal path As String) As String
    Dim wb As Workbook, ws As Worksheet, c As Object
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Worksheets(1)
    ws.Name = "One"
    ws.Range("A1:C1").Value = Array("", "Rev", "Cost")
    ws.Range("A2:C2").Value = Array("a", 10, 4)
    ws.Range("A3:C3").Value = Array("b", 20, 6)
    ws.Range("E1:F1").Value = Array("x", "y")
    Set c = ws.ChartObjects.Add(48, 96, 240, 144).Chart
    c.SetSourceData ws.Range("A1:C3")
    c.SeriesCollection(2).XValues = "=One!$E$1:$F$1"
    wb.SaveAs path, xlOpenXMLWorkbook
    wb.Close False
End Function

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


def chart_part(path: Path) -> str:
    with zipfile.ZipFile(path) as package:
        return package.read('xl/charts/chart1.xml').decode('utf-8')


def attempt(code: str, macro: str, arguments: tuple[object, ...]) -> str:
    """One run in a new Excel session, again once if the session fails to start or the run fails."""
    answer = ''
    for _turn in range(2):
        try:
            with ExcelSession(HarnessConfig(lock_wait_s=600)) as excel:
                done = excel.run_vba(code, macro, arguments, timeout=120.0, line_numbers=False)
        except HarnessError as failed:
            answer = f'harness: {failed}'
            continue
        answer = str(done.value) if done.outcome == 'passed' else f'{done.outcome}: {done.message}'
        if done.outcome == 'passed':
            break
    return answer


def main() -> None:
    folder = Path(tempfile.mkdtemp())
    code = source()
    made = folder / OWN
    answer = attempt(code, 'Make', (str(made),))
    assert made.exists(), answer
    without_folder(made, FIXTURES / OWN)
    originals = {name: chart_part(FIXTURES / name) for name in BOOKS}
    cases: list[list[object]] = []
    for name in BOOKS:
        for which, (label, _) in enumerate(EDITS):
            copy, out = folder / 'in.xlsx', folder / f'{name}{which}.xlsx'
            shutil.copy(FIXTURES / name, copy)
            answer = attempt(code, 'Act', (str(copy), str(out), which))
            cases.append([name, which, answer, chart_part(out) if out.exists() else ''])
            print(name, label, '=>', answer, flush=True)
    record = {'measured_at': datetime.datetime.now(datetime.timezone.utc).isoformat(), 'source': code, 'edits': EDITS,
              'originals': originals, 'cases': cases}
    (FIXTURES / 'loaded_chart_formula_categories.json').write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')


if __name__ == '__main__':
    main()
