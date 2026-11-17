"""Measure Chart.SaveChartTemplate, ApplyChartTemplate, SetBackgroundPicture and Paste: the .crtx package Excel writes,
and the chart part a workbook saves after each edit, on a plain embedded chart. The template comes from a styled line
chart; the picture is a small PNG this script writes. Files go under TEMP alone. Needs Excel:

    python scripts/measure_chart_templates.py

writes tests/fixtures/chart_templates.json.
"""

from __future__ import annotations

import base64
import datetime
import json
import os
import tempfile
import zipfile
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig
from pyvbaharness.results import HarnessError

FIXTURES = Path(__file__).resolve().parents[1] / 'tests/fixtures'
TEMP = Path(os.environ['TEMP'])
TEMPLATE = TEMP / 'pyov_chart_template.crtx'
PICTURE = TEMP / 'pyov_chart_background.png'
#: A red 2x2 PNG.
PNG = base64.b64decode('iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAIAAAD91JpzAAAAEklEQVR4nGP4z8DAwMDAxAAAAP8KAgBHZtmLAAAAAElFTkSuQmCC')
C, W = 'c', 'c.Parent.Parent'
#: Each: what it is, and the statements, with c the chart, t the template's path and p the picture's. A read follows in r.
EDITS = (
    ('nothing', ''),
    ('save template', f'{C}.SaveChartTemplate t & ".2.crtx": r = (Dir(t & ".2.crtx") <> ""): Kill t & ".2.crtx"'),
    ('apply template', f'{C}.ApplyChartTemplate t: r = {C}.ChartType'),
    ('apply missing', f'{C}.ApplyChartTemplate t & ".none.crtx": r = {C}.ChartType'),
    ('background', f'{C}.SetBackgroundPicture p: r = {C}.ChartArea.Format.Fill.Type'),
    ('background missing', f'{C}.SetBackgroundPicture p & ".none.png": r = {C}.ChartArea.Format.Fill.Type'),
    ('paste row', f'{W}.Range("A4:C4").Copy: {C}.Paste: r = {C}.SeriesCollection.Count'),
    ('paste rows', f'{W}.Range("A4:C5").Copy: {C}.Paste: r = {C}.SeriesCollection.Count'),
    ('paste formats', f'{W}.ChartObjects(2).Chart.ChartArea.Copy: {C}.Paste xlFormats: r = {C}.ChartType'),
    ('paste nothing', f'Application.CutCopyMode = False: {C}.Paste: r = {C}.SeriesCollection.Count'),
)


def source() -> str:
    lines = ['Private Sub Edit(ByVal which As Long, c As Object, ByVal t As String, ByVal p As String, r As Variant)',
             '    Select Case which']
    lines += [f'    Case {index}: {statements}' for index, (_, statements) in enumerate(EDITS) if statements]
    lines += ['    End Select', 'End Sub', r'''
Private Function Book() As Workbook
    Dim wb As Workbook, ws As Worksheet
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Worksheets(1)
    ws.Name = "One"
    ws.Range("A1:C1").Value = Array("", "Rev", "Cost")
    ws.Range("A2:C2").Value = Array("a", 10, 4)
    ws.Range("A3:C3").Value = Array("b", 20, 6)
    ws.Range("A4:C4").Value = Array("c", 30, 8)
    ws.Range("A5:C5").Value = Array("d", 40, 9)
    ws.ChartObjects.Add(48, 96, 240, 144).Chart.SetSourceData ws.Range("A1:C3")
    With ws.Shapes.AddChart2(227, xlLine, 320, 96, 240, 144).Chart
        .SetSourceData ws.Range("A1:C3")
    End With
    Set Book = wb
End Function

Public Function MakeTemplate(ByVal t As String) As String
    Dim wb As Workbook
    Set wb = Book()
    wb.Worksheets(1).ChartObjects(2).Chart.SaveChartTemplate t
    wb.Close False
    MakeTemplate = CStr(Dir(t) <> "")
End Function

Public Function Act(ByVal out As String, ByVal which As Long, ByVal t As String, ByVal p As String) As String
    Dim wb As Workbook, r As Variant, s As String
    Set wb = Book()
    On Error Resume Next
    Edit which, wb.Worksheets(1).ChartObjects(1).Chart, t, p, r
    s = Err.Number & " " & TypeName(r) & " " & r
    On Error GoTo 0
    Application.CutCopyMode = False
    wb.SaveAs out, xlOpenXMLWorkbook
    wb.Close False
    Act = s
End Function
''']
    return '\n'.join(lines)


def parts(path: Path) -> dict[str, str]:
    """A package's parts: XML and relationships as text, anything else as base64."""
    with zipfile.ZipFile(path) as package:
        found: dict[str, str] = {}
        for name in package.namelist():
            data = package.read(name)
            found[name] = data.decode('utf-8') if name.endswith(('.xml', '.rels')) else 'base64:' + base64.b64encode(data).decode()
        return found


def run(code: str, macro: str, arguments: tuple[object, ...]) -> str:
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
    PICTURE.write_bytes(PNG)
    TEMPLATE.unlink(missing_ok=True)
    made = run(code, 'MakeTemplate', (str(TEMPLATE),))
    assert TEMPLATE.exists(), made
    template = parts(TEMPLATE)
    cases: list[list[object]] = []
    for which, (label, _) in enumerate(EDITS):
        out = folder / f'{which}.xlsx'
        answer = run(code, 'Act', (str(out), which, str(TEMPLATE), str(PICTURE)))
        saved = {name: text for name, text in parts(out).items() if name.startswith(('xl/charts/', 'xl/media/'))} if out.exists() else {}
        cases.append([which, answer, saved])
        print(label, '=>', answer, flush=True)
    record = {'measured_at': datetime.datetime.now(datetime.timezone.utc).isoformat(), 'source': code, 'edits': EDITS,
              'picture': base64.b64encode(PNG).decode(), 'template': template, 'cases': cases}
    (FIXTURES / 'chart_templates.json').write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    TEMPLATE.unlink(missing_ok=True)
    PICTURE.unlink(missing_ok=True)


if __name__ == '__main__':
    main()
