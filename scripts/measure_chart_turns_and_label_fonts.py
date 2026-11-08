"""What Excel saves for three loaded-chart gaps, on a new chart and on the plain loaded book: a typed point label given a font, a
turn to an area, pie or doughnut chart, and one series turned to a line. Each case runs in a new Excel session. Needs Excel:

    python scripts/measure_chart_turns_and_label_fonts.py

writes tests/fixtures/chart_turns_and_label_fonts.json.
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
LOADED = 'loaded_chart_plain.xlsx'
P = 'c.SeriesCollection(1).Points(1)'
#: Each: what it is, and the statements, with c the chart. A read follows in r, for its answer and type.
EDITS = (
    ('nothing', ''),
    ('label text then font', f'{P}.HasDataLabel = True: {P}.DataLabel.Text = "P": {P}.DataLabel.Font.Size = 12: r = {P}.DataLabel.Font.Size'),
    ('label font then text', f'{P}.HasDataLabel = True: {P}.DataLabel.Font.Size = 12: {P}.DataLabel.Text = "P": r = {P}.DataLabel.Text'),
    ('label font only', f'{P}.HasDataLabel = True: {P}.DataLabel.Font.Size = 12: r = {P}.DataLabel.Font.Size'),
    ('label text then bold', f'{P}.HasDataLabel = True: {P}.DataLabel.Text = "P": {P}.DataLabel.Font.Bold = True: r = {P}.DataLabel.Font.Bold'),
    ('type area', 'c.ChartType = xlArea: r = c.ChartType'),
    ('type area stacked', 'c.ChartType = xlAreaStacked: r = c.ChartType'),
    ('type pie', 'c.ChartType = xlPie: r = c.ChartType'),
    ('type doughnut', 'c.ChartType = xlDoughnut: r = c.ChartType'),
    ('type area and back', 'c.ChartType = xlArea: c.ChartType = xlColumnClustered: r = c.ChartType'),
    ('series type line', 'c.SeriesCollection(2).ChartType = xlLine: r = c.SeriesCollection(2).ChartType'),
    ('series type line markers', 'c.SeriesCollection(2).ChartType = xlLineMarkers: r = c.SeriesCollection(2).ChartType'),
)


def source() -> str:
    lines = ['Private Sub Edit(ByVal which As Long, c As Object, r As Variant)', '    Select Case which']
    lines += [f'    Case {index}: {statements}' for index, (_, statements) in enumerate(EDITS) if statements]
    lines += ['    End Select', 'End Sub', r'''
Private Function Finish(wb As Workbook, c As Object, ByVal out As String, ByVal which As Long) As String
    Dim r As Variant, s As String
    On Error Resume Next
    Edit which, c, r
    s = Err.Number & " " & TypeName(r) & " " & r
    On Error GoTo 0
    Application.DisplayAlerts = False
    wb.SaveAs out, xlOpenXMLWorkbook
    Application.DisplayAlerts = True
    wb.Close False
    Finish = s
End Function

Public Function Loaded(ByVal path As String, ByVal out As String, ByVal which As Long) As String
    Dim wb As Workbook
    Set wb = Workbooks.Open(path)
    Loaded = Finish(wb, wb.Worksheets(1).ChartObjects(1).Chart, out, which)
End Function

Public Function Fresh(ByVal out As String, ByVal which As Long) As String
    Dim wb As Workbook, ws As Worksheet, c As Object
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Worksheets(1)
    ws.Name = "One"
    ws.Range("A1:C1").Value = Array("", "Rev", "Cost")
    ws.Range("A2:C2").Value = Array("a", 10, 4)
    ws.Range("A3:C3").Value = Array("b", 20, 6)
    Set c = ws.ChartObjects.Add(48, 96, 240, 144).Chart
    c.SetSourceData ws.Range("A1:C3")
    Fresh = Finish(wb, c, out, which)
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
    for where in ('new', LOADED):
        for which, (label, _) in enumerate(EDITS):
            answer, saved = '', {}
            for _turn in range(3):
                copy, out = folder / 'in.xlsx', folder / f'{where}{which}.xlsx'
                shutil.copy(FIXTURES / LOADED, copy)
                macro, arguments = ('Fresh', (str(out), which)) if where == 'new' else ('Loaded', (str(copy), str(out), which))
                try:
                    with ExcelSession(HarnessConfig(lock_wait_s=600)) as excel:
                        done = excel.run_vba(code, macro, arguments, timeout=120.0, line_numbers=False)
                except HarnessError as failed:
                    # Excel's worker can be slow to start; the case runs again in a new session.
                    answer = f'harness: {failed}'
                    continue
                answer = done.value if done.outcome == 'passed' else f'{done.outcome}: {done.message}'
                if done.outcome == 'passed':
                    saved = chart_parts(out)
                    break
            cases.append([where, which, answer, saved])
            print(where, label, '=>', answer, flush=True)
    record = {'measured_at': datetime.datetime.now(datetime.timezone.utc).isoformat(), 'source': code, 'edits': EDITS,
              'original': chart_parts(FIXTURES / LOADED), 'cases': cases}
    (FIXTURES / 'chart_turns_and_label_fonts.json').write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8', newline='\n')


if __name__ == '__main__':
    main()
