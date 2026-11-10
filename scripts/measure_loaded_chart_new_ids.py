"""The unique id numbers Excel gives what it adds to a loaded chart: a point label, a point fill and a new series, on
books with trendlines and with a manual layout. Needs Excel:

    python scripts/measure_loaded_chart_new_ids.py

writes tests/fixtures/loaded_chart_new_ids.json and the books beside it.
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
S1, S2 = 'c.SeriesCollection(1)', 'c.SeriesCollection(2)'
#: Each book: its name, and what the chart takes before it is saved.
BOOKS = (
    ('loaded_chart_trendline.xlsx', f'{S2}.Trendlines.Add'),
    ('loaded_chart_trendlines_both.xlsx', f'{S1}.Trendlines.Add: {S2}.Trendlines.Add'),
    ('loaded_chart_trendlines_two.xlsx', f'{S1}.Trendlines.Add: {S1}.Trendlines.Add'),
    ('loaded_chart_layout.xlsx', 'c.Legend.Left = 10'),
)
#: Each: what it is, and the statements, with c the chart. A read follows in r, for its answer and type.
EDITS = (
    ('nothing', ''),
    ('point label', f'{S1}.Points(1).HasDataLabel = True: r = {S1}.Points(1).HasDataLabel'),
    ('point fill', f'{S1}.Points(1).Format.Fill.ForeColor.RGB = 255: r = {S1}.Points(1).Format.Fill.ForeColor.RGB'),
    ('new series', 'c.SeriesCollection.NewSeries: r = c.SeriesCollection.Count'),
)


def source() -> str:
    lines = ['Private Sub Dress(ByVal which As Long, c As Object)', '    Select Case which']
    lines += [f'    Case {index}: {statements}' for index, (_, statements) in enumerate(BOOKS)]
    lines += ['    End Select', 'End Sub', 'Private Sub Edit(ByVal which As Long, c As Object, r As Variant)', '    Select Case which']
    lines += [f'    Case {index}: {statements}' for index, (_, statements) in enumerate(EDITS) if statements]
    lines += ['    End Select', 'End Sub', r'''
Public Function Make(ByVal path As String, ByVal which As Long) As String
    Dim wb As Workbook, ws As Worksheet, c As Object
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Worksheets(1)
    ws.Name = "One"
    ws.Range("A1:C1").Value = Array("", "Rev", "Cost")
    ws.Range("A2:C2").Value = Array("a", 10, 4)
    ws.Range("A3:C3").Value = Array("b", 20, 6)
    Set c = ws.ChartObjects.Add(48, 96, 240, 144).Chart
    c.SetSourceData ws.Range("A1:C3")
    Dress which, c
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
    originals: dict[str, str] = {}
    cases: list[list[object]] = []
    for book, (name, _) in enumerate(BOOKS):
        made = folder / name
        answer = attempt(code, 'Make', (str(made), book))
        assert made.exists(), answer
        without_folder(made, FIXTURES / name)
        originals[name] = chart_part(FIXTURES / name)
        for which, (label, _) in enumerate(EDITS):
            copy, out = folder / 'in.xlsx', folder / f'{book}-{which}.xlsx'
            shutil.copy(FIXTURES / name, copy)
            answer = attempt(code, 'Act', (str(copy), str(out), which))
            cases.append([name, which, answer, chart_part(out) if out.exists() else ''])
            print(name, label, '=>', answer, flush=True)
    record = {'measured_at': datetime.datetime.now(datetime.timezone.utc).isoformat(), 'source': code, 'books': BOOKS,
              'edits': EDITS, 'originals': originals, 'cases': cases}
    (FIXTURES / 'loaded_chart_new_ids.json').write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')


if __name__ == '__main__':
    main()
