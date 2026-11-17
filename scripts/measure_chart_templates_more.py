"""Measure Chart.SaveChartTemplate from four kinds of chart, and Chart.ApplyChartTemplate of each template onto three kinds
of chart: the .crtx packages Excel writes, and every chart, theme, relationship and content-type part a workbook saves
after an apply. Files go under TEMP alone. Needs Excel:

    python scripts/measure_chart_templates_more.py

writes tests/fixtures/chart_templates_more.json.
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
#: Each template's source: what it is, and the statements that make it, with ws the sheet, giving c the chart.
SOURCES = (
    ('line 227', 'Set c = ws.Shapes.AddChart2(227, xlLine, 320, 96, 240, 144).Chart: c.SetSourceData ws.Range("A1:C3")'),
    ('column 201', 'Set c = ws.Shapes.AddChart2(201, xlColumnClustered, 320, 96, 240, 144).Chart: c.SetSourceData ws.Range("A1:C3")'),
    ('plain', 'Set c = ws.ChartObjects.Add(320, 96, 240, 144).Chart: c.SetSourceData ws.Range("A1:C3")'),
    ('column 201 dressed', 'Set c = ws.Shapes.AddChart2(201, xlColumnClustered, 320, 96, 240, 144).Chart: c.SetSourceData ws.Range("A1:C3"): '
                           'c.HasTitle = True: c.ChartTitle.Text = "T": c.SeriesCollection(1).Format.Fill.ForeColor.RGB = 255'),
)
#: Each target: what it is, and the statements that make it, giving c the chart a template is applied to.
TARGETS = (
    ('plain two', 'Set c = ws.ChartObjects.Add(48, 96, 240, 144).Chart: c.SetSourceData ws.Range("A1:C3")'),
    ('column 201 two', 'Set c = ws.Shapes.AddChart2(201, xlColumnClustered, 48, 96, 240, 144).Chart: c.SetSourceData ws.Range("A1:C3")'),
    ('plain three', 'Set c = ws.ChartObjects.Add(48, 96, 240, 144).Chart: c.SetSourceData ws.Range("A1:C4")'),
)


def source() -> str:
    lines = ['Private Function Made(ws As Worksheet, ByVal which As Long, ByVal target As Boolean) As Object', '    Dim c As Object']
    lines += ['    If target Then', '        Select Case which']
    lines += [f'        Case {index}: {statements}' for index, (_, statements) in enumerate(TARGETS)]
    lines += ['        End Select', '    Else', '        Select Case which']
    lines += [f'        Case {index}: {statements}' for index, (_, statements) in enumerate(SOURCES)]
    lines += ['        End Select', '    End If', '    Set Made = c', 'End Function', r'''
Private Function Book() As Workbook
    Dim wb As Workbook, ws As Worksheet
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Worksheets(1)
    ws.Name = "One"
    ws.Range("A1:C1").Value = Array("", "Rev", "Cost")
    ws.Range("A2:C2").Value = Array("a", 10, 4)
    ws.Range("A3:C3").Value = Array("b", 20, 6)
    ws.Range("A4:C4").Value = Array("c", 30, 8)
    Set Book = wb
End Function

Public Function MakeTemplate(ByVal which As Long, ByVal t As String) As String
    Dim wb As Workbook
    Set wb = Book()
    Made(wb.Worksheets(1), which, False).SaveChartTemplate t
    wb.Close False
    MakeTemplate = CStr(Dir(t) <> "")
End Function

Public Function Apply(ByVal out As String, ByVal which As Long, ByVal t As String) As String
    Dim wb As Workbook, c As Object, s As String
    Set wb = Book()
    Set c = Made(wb.Worksheets(1), which, True)
    On Error Resume Next
    c.ApplyChartTemplate t
    s = Err.Number & " " & c.ChartType & " " & c.SeriesCollection.Count
    On Error GoTo 0
    wb.SaveAs out, xlOpenXMLWorkbook
    wb.Close False
    Apply = s
End Function
''']
    return '\n'.join(lines)


def parts(path: Path, prefixes: tuple[str, ...] = ('',)) -> dict[str, str]:
    """A package's parts under these prefixes: XML and relationships as text, anything else as base64."""
    with zipfile.ZipFile(path) as package:
        found: dict[str, str] = {}
        for name in package.namelist():
            if not name.startswith(prefixes):
                continue
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
    templates: list[list[object]] = []
    applied: list[list[object]] = []
    for which, (label, _) in enumerate(SOURCES):
        path = TEMP / f'pyov_template_{which}.crtx'
        path.unlink(missing_ok=True)
        answer = run(code, 'MakeTemplate', (which, str(path)))
        templates.append([which, answer, parts(path) if path.exists() else {}])
        print('template', label, '=>', answer, flush=True)
        for target, (named, _) in enumerate(TARGETS):
            out = folder / f'{which}-{target}.xlsx'
            answer = run(code, 'Apply', (str(out), target, str(path)))
            saved = parts(out, ('xl/charts/', 'xl/theme/', 'xl/media/', '[Content_Types].xml')) if out.exists() else {}
            applied.append([which, target, answer, saved])
            print('  apply to', named, '=>', answer, flush=True)
        path.unlink(missing_ok=True)
    record = {'measured_at': datetime.datetime.now(datetime.timezone.utc).isoformat(), 'source': code, 'sources': SOURCES,
              'targets': TARGETS, 'templates': templates, 'applied': applied}
    (FIXTURES / 'chart_templates_more.json').write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')


if __name__ == '__main__':
    main()
