"""What Excel's save does to a loaded chart's part for one fill or line edit at a time, and what it answers after. On four books Excel
saved: the plain and dressed books of scripts/measure_loaded_chart_edits.py, the labelled book of scripts/measure_loaded_chart_labels.py
and the themed book of scripts/measure_loaded_chart_looks.py. Each case opens its book in a new Excel session, as the model opens a file.
Needs Excel:

    python scripts/measure_loaded_chart_look_edits.py

writes tests/fixtures/loaded_chart_look_edits.json, which tests/test_excel_loaded_chart_look_edits.py replays.
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
BOOKS = ('loaded_chart_plain.xlsx', 'loaded_chart_dressed.xlsx', 'loaded_chart_labelled.xlsx', 'loaded_chart_themed.xlsx')
S1, S2 = 'c.SeriesCollection(1)', 'c.SeriesCollection(2)'
F1, L1 = f'{S1}.Format.Fill', f'{S1}.Format.Line'
P1 = f'{S1}.Points(1).Format'
#: Each: what it is, and the statements, with c the chart. A read follows in r, for its answer and type.
EDITS = (
    ('nothing', ''),
    ('fill colour', f'{F1}.ForeColor.RGB = 255: r = {F1}.ForeColor.RGB'),
    ('fill colour second', f'{S2}.Format.Fill.ForeColor.RGB = 65280: r = {S2}.Format.Fill.ForeColor.RGB'),
    ('fill colour twice', f'{F1}.ForeColor.RGB = 255: {F1}.ForeColor.RGB = 65280: r = {F1}.ForeColor.RGB'),
    ('fill hidden', f'{F1}.Visible = msoFalse: r = {F1}.Visible'),
    ('fill colour transparency', f'{F1}.ForeColor.RGB = 255: {F1}.Transparency = 0.5: r = {F1}.Transparency'),
    ('fill solid', f'{F1}.Solid: r = {F1}.Type'),
    ('line colour', f'{L1}.ForeColor.RGB = 255: r = {L1}.ForeColor.RGB'),
    ('line weight', f'{L1}.Weight = 2: r = {L1}.Weight'),
    ('line colour weight', f'{L1}.ForeColor.RGB = 255: {L1}.Weight = 2: r = {L1}.Weight'),
    ('line hidden', f'{L1}.Visible = msoFalse: r = {L1}.Visible'),
    ('line dash', f'{L1}.ForeColor.RGB = 255: {L1}.DashStyle = msoLineDash: r = {L1}.DashStyle'),
    ('interior colour', f'{S1}.Interior.Color = 255: r = {S1}.Interior.Color'),
    ('fill and line', f'{F1}.ForeColor.RGB = 255: {L1}.ForeColor.RGB = 65280: r = {L1}.ForeColor.RGB'),
    ('point fill', f'{P1}.Fill.ForeColor.RGB = 255: r = {P1}.Fill.ForeColor.RGB'),
    ('point fill second point', f'{S1}.Points(2).Format.Fill.ForeColor.RGB = 255: r = {S1}.Points(2).Format.Fill.ForeColor.RGB'),
    ('point line', f'{P1}.Line.ForeColor.RGB = 255: r = {P1}.Line.ForeColor.RGB'),
    ('point fill then series fill', f'{P1}.Fill.ForeColor.RGB = 255: {F1}.ForeColor.RGB = 65280: r = {P1}.Fill.ForeColor.RGB'),
    ('two point fills', f'{S1}.Points(2).Format.Fill.ForeColor.RGB = 255: {P1}.Fill.ForeColor.RGB = 65280: r = {P1}.Fill.ForeColor.RGB'),
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
    (FIXTURES / 'loaded_chart_look_edits.json').write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8', newline='\n')


if __name__ == '__main__':
    main()
