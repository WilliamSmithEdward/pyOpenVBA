"""What Excel's save does to a loaded chart's part for one data label edit at a time, and what it answers after. On three books Excel
saved: the plain and dressed books of scripts/measure_loaded_chart_edits.py and the labelled book of
scripts/measure_loaded_chart_labels.py. Each case opens its book in a new Excel session, as the model opens a file. Needs Excel:

    python scripts/measure_loaded_chart_label_edits.py

writes tests/fixtures/loaded_chart_label_edits.json, which tests/test_excel_loaded_chart_label_edits.py replays.
"""

from __future__ import annotations

import datetime
import json
import shutil
import tempfile
import zipfile
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

FIXTURES = Path(__file__).resolve().parents[1] / 'tests/fixtures'
BOOKS = ('loaded_chart_plain.xlsx', 'loaded_chart_dressed.xlsx', 'loaded_chart_labelled.xlsx')
S1, S2 = 'c.SeriesCollection(1)', 'c.SeriesCollection(2)'
#: Each: what it is, and the statements, with c the chart. A read follows in r, for its answer and type.
EDITS = (
    ('nothing', ''),
    ('labels on', f'{S1}.HasDataLabels = True: r = {S1}.HasDataLabels'),
    ('labels on second', f'{S2}.HasDataLabels = True: r = {S2}.HasDataLabels'),
    ('labels off', f'{S1}.HasDataLabels = False: r = {S1}.HasDataLabels'),
    ('labels on then off', f'{S1}.HasDataLabels = True: {S1}.HasDataLabels = False: r = {S1}.HasDataLabels'),
    ('labels on twice', f'{S1}.HasDataLabels = True: {S1}.HasDataLabels = True: r = {S1}.HasDataLabels'),
    ('point label on', f'{S1}.Points(1).HasDataLabel = True: r = {S1}.Points(1).HasDataLabel'),
    ('point label off', f'{S1}.Points(1).HasDataLabel = False: r = {S1}.Points(1).HasDataLabel'),
    ('point label text', f'{S1}.Points(1).HasDataLabel = True: {S1}.Points(1).DataLabel.Text = "P": r = {S1}.Points(1).DataLabel.Text'),
    ('point label text second', f'{S2}.Points(1).HasDataLabel = True: {S2}.Points(1).DataLabel.Text = "P": r = {S2}.Points(1).DataLabel.Text'),
    ('two point labels', f'{S1}.Points(2).HasDataLabel = True: {S1}.Points(2).DataLabel.Text = "Q": '
                         f'{S1}.Points(1).HasDataLabel = True: {S1}.Points(1).DataLabel.Text = "P": r = {S1}.Points(2).DataLabel.Text'),
    ('point labels on both', f'{S1}.Points(1).HasDataLabel = True: {S1}.Points(1).DataLabel.Text = "P": '
                             f'{S2}.Points(1).HasDataLabel = True: {S2}.Points(1).DataLabel.Text = "R": r = {S2}.Points(1).DataLabel.Text'),
    ('point text second point', f'{S2}.Points(2).HasDataLabel = True: {S2}.Points(2).DataLabel.Text = "Z": r = {S2}.Points(2).DataLabel.Text'),
    ('point label then labels off', f'{S1}.Points(1).HasDataLabel = True: {S1}.Points(1).DataLabel.Text = "P": '
                                    f'{S1}.HasDataLabels = False: r = {S1}.Points(1).HasDataLabel'),
    ('show category', f'{S1}.HasDataLabels = True: {S1}.DataLabels.ShowCategoryName = True: r = {S1}.DataLabels.ShowCategoryName'),
    ('show series name', f'{S1}.HasDataLabels = True: {S1}.DataLabels.ShowSeriesName = True: r = {S1}.DataLabels.ShowSeriesName'),
    ('show value off', f'{S1}.HasDataLabels = True: {S1}.DataLabels.ShowCategoryName = True: {S1}.DataLabels.ShowValue = False: '
                       f'r = {S1}.DataLabels.ShowValue'),
    ('position', f'{S1}.HasDataLabels = True: {S1}.DataLabels.Position = xlLabelPositionInsideEnd: r = {S1}.DataLabels.Position'),
    ('number format', f'{S1}.HasDataLabels = True: {S1}.DataLabels.NumberFormat = "0.0": r = {S1}.DataLabels.NumberFormat'),
    ('separator', f'{S1}.HasDataLabels = True: {S1}.DataLabels.ShowCategoryName = True: {S1}.DataLabels.Separator = "; ": '
                  f'r = {S1}.DataLabels.Separator'),
    ('point show category', f'{S1}.Points(1).HasDataLabel = True: {S1}.Points(1).DataLabel.ShowCategoryName = True: '
                            f'r = {S1}.Points(1).DataLabel.ShowCategoryName'),
    ('point position', f'{S1}.Points(1).HasDataLabel = True: {S1}.Points(1).DataLabel.Position = xlLabelPositionInsideBase: '
                       f'r = {S1}.Points(1).DataLabel.Position'),
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
            for _turn in range(2):
                copy, out = folder / 'in.xlsx', folder / f'{name}{which}.xlsx'
                shutil.copy(FIXTURES / name, copy)
                with ExcelSession(HarnessConfig(lock_wait_s=600)) as excel:
                    done = excel.run_vba(code, 'Act', (str(copy), str(out), which), timeout=120.0, line_numbers=False)
                answer = done.value if done.outcome == 'passed' else f'{done.outcome}: {done.message}'
                if done.outcome == 'passed':
                    saved = chart_parts(out)
                    break
            cases.append([name, which, answer, saved])
            print(name, label, '=>', answer, flush=True)
    record = {'measured_at': datetime.datetime.now(datetime.timezone.utc).isoformat(), 'source': code, 'edits': EDITS, 'originals': originals,
              'cases': cases}
    (FIXTURES / 'loaded_chart_label_edits.json').write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8', newline='\n')


if __name__ == '__main__':
    main()
