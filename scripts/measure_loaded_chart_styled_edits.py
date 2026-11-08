"""What Excel answers and saves for each loaded-chart edit already measured on the plain book, on a chart with a style: one
made in the session with Shapes.AddChart2(201) and nothing else, and the same chart saved and opened again. Where Excel's
save of the loaded chart is its save of the new one, the model's new styled chart predicts it. Needs Excel:

    python scripts/measure_loaded_chart_styled_edits.py

writes tests/fixtures/loaded_chart_styled_edits.json and tests/fixtures/loaded_chart_styled.xlsx.
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
import measure_chart_turns_and_label_fonts as turns  # noqa: E402
import measure_loaded_chart_edits as elements  # noqa: E402
import measure_loaded_chart_label_edits as labels  # noqa: E402
import measure_loaded_chart_look_edits as looks  # noqa: E402
import measure_loaded_chart_structure_edits as structure  # noqa: E402

FIXTURES = Path(__file__).resolve().parents[1] / 'tests/fixtures'
BOOK = 'loaded_chart_styled.xlsx'
#: Each edit by the list it comes from, with c the chart and co its object. A read follows in r, for its answer and type.
EDITS = tuple([('nothing', '')] + [
    (f'{group} {label}', statements)
    for group, edits in (('element', elements.EDITS), ('look', looks.EDITS), ('labels', labels.EDITS),
                         ('structure', structure.EDITS), ('turns', turns.EDITS))
    for label, statements in edits if statements])


def source() -> str:
    lines = ['Private Sub Edit(ByVal which As Long, c As Object, co As Object, r As Variant)', '    Select Case which']
    lines += [f'    Case {index}: {statements}' for index, (_, statements) in enumerate(EDITS) if statements]
    lines += ['    End Select', 'End Sub', r'''
Private Function Made(ws As Worksheet) As Object
    ws.Name = "One"
    ws.Range("A1:C1").Value = Array("", "Rev", "Cost")
    ws.Range("A2:C2").Value = Array("a", 10, 4)
    ws.Range("A3:C3").Value = Array("b", 20, 6)
    Set Made = ws.Shapes.AddChart2(201, xlColumnClustered, 48, 96, 240, 144)
    Made.Chart.SetSourceData ws.Range("A1:C3")
End Function

Private Function Finish(wb As Workbook, co As Object, ByVal out As String, ByVal which As Long) As String
    Dim r As Variant, s As String
    On Error Resume Next
    Edit which, co.Chart, co, r
    s = Err.Number & " " & TypeName(r) & " " & r
    On Error GoTo 0
    Application.DisplayAlerts = False
    wb.SaveAs out, xlOpenXMLWorkbook
    Application.DisplayAlerts = True
    wb.Close False
    Finish = s
End Function

Public Function Make(ByVal path As String) As String
    Dim wb As Workbook
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Made wb.Worksheets(1)
    wb.SaveAs path, xlOpenXMLWorkbook
    wb.Close False
End Function

Public Function Fresh(ByVal out As String, ByVal which As Long) As String
    Dim wb As Workbook
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Made wb.Worksheets(1)
    Fresh = Finish(wb, wb.Worksheets(1).ChartObjects(1), out, which)
End Function

Public Function Loaded(ByVal path As String, ByVal out As String, ByVal which As Long) As String
    Dim wb As Workbook
    Set wb = Workbooks.Open(path)
    Loaded = Finish(wb, wb.Worksheets(1).ChartObjects(1), out, which)
End Function
''']
    return '\n'.join(lines)


def chart_parts(path: Path) -> dict[str, str]:
    with zipfile.ZipFile(path) as package:
        return {name: package.read(name).decode('utf-8') for name in package.namelist()
                if name.startswith('xl/charts/') and name.endswith(('.xml', '.rels'))}


def attempt(code: str, macro: str, arguments: tuple[object, ...], out: Path) -> tuple[str, dict[str, str]]:
    """One case in a new Excel session, again once if the session fails to start or the edit kills Excel."""
    answer = ''
    for _turn in range(2):
        try:
            with ExcelSession(HarnessConfig(lock_wait_s=600)) as excel:
                done = excel.run_vba(code, macro, arguments, timeout=120.0, line_numbers=False)
        except HarnessError as failed:
            answer = f'harness: {failed}'
            continue
        answer = done.value if done.outcome == 'passed' else f'{done.outcome}: {done.message}'
        if done.outcome == 'passed':
            return str(answer), chart_parts(out)
    return answer, {}


def main() -> None:
    folder = Path(tempfile.mkdtemp())
    code = source()
    made = folder / BOOK
    answer, _ = attempt(code, 'Make', (str(made),), made)
    assert made.exists(), answer
    elements.without_folder(made, FIXTURES / BOOK)
    cases: list[list[object]] = []
    for which, (label, _) in enumerate(EDITS):
        out = folder / f'new{which}.xlsx'
        cases.append(['new', which, *attempt(code, 'Fresh', (str(out), which), out)])
        copy, out = folder / 'in.xlsx', folder / f'loaded{which}.xlsx'
        shutil.copy(FIXTURES / BOOK, copy)
        cases.append(['loaded', which, *attempt(code, 'Loaded', (str(copy), str(out), which), out)])
        print(label, '=>', cases[-2][2], '|', cases[-1][2], flush=True)
    original = chart_parts(FIXTURES / BOOK)
    for case in cases:
        # A save's chart part, and its other parts where they are not the original's: its style and colours stay as they are.
        case[3] = {name: text for name, text in case[3].items() if name == 'xl/charts/chart1.xml' or original.get(name) != text}
    record = {'measured_at': datetime.datetime.now(datetime.timezone.utc).isoformat(), 'source': code, 'edits': EDITS,
              'original': original, 'cases': cases}
    (FIXTURES / 'loaded_chart_styled_edits.json').write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')


if __name__ == '__main__':
    main()
