"""A chart read from a file, when the cells it reads are moved or hidden: what it reads afterwards, and each reference and
cached value in the chart part Excel saves.

It builds a workbook with this project, with a chart on Data and one on Other. Live Excel opens it afresh for each case,
runs the case's statements, reads READS and saves it.

    python scripts/measure_loaded_chart_references.py

writes tests/fixtures/loaded_chart_references.json, which tests/test_excel_loaded_chart_references.py replays.
"""
from __future__ import annotations

import json
import re
import sys
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / 'tests/fixtures/loaded_chart_references.json'
BUILD = '''Public Sub Build()
    Dim d As Worksheet, o As Worksheet, ch As Object
    Set d = ActiveWorkbook.Worksheets(1): d.Name = "Data"
    Set o = ActiveWorkbook.Worksheets.Add(After:=d): o.Name = "Other": d.Activate
    d.Range("A1:A5").Value = Application.Transpose(Array("k", "a", "b", "c", "d"))
    d.Range("B1:B5").Value = Application.Transpose(Array("v", 1, 2, 3, 4))
    d.Range("C1:C5").Value = Application.Transpose(Array("w", 5, 6, 7, 8))
    Set ch = d.ChartObjects.Add(300, 100, 300, 200).Chart
    ch.SetSourceData d.Range("A1:C5"), xlColumns
    ch.HasTitle = True
    ch.ChartTitle.Formula = "=Data!$A$1"
    Set ch = o.ChartObjects.Add(300, 100, 300, 200).Chart
    ch.SetSourceData d.Range("A1:B5"), xlColumns
End Sub
'''
EDITS = ['', 'd.Rows(1).Insert', 'd.Rows(3).Insert', 'd.Rows(9).Insert', 'd.Rows(3).Delete', 'd.Rows("3:4").Delete', 'd.Columns(1).Insert', 'd.Columns(3).Insert', 'd.Columns(5).Insert', 'd.Range("B3").Insert xlShiftDown', 'd.Range("B3:C3").Delete xlShiftUp', 'd.Range("B2:B5").Insert xlShiftToRight',
         'd.Range("A1:C5").Cut d.Range("E7")', 'd.Range("B2:B5").Cut d.Range("E7")', 'd.Name = "Moved"', 'd.Name = "With Space"', 'o.Name = "Moved"', 'o.Rows(1).Insert', 'd.Rows(1).Insert: d.Columns(1).Insert', 'd.Rows(3).Hidden = True', 'd.Range("A1:C5").AutoFilter 2, ">2"', 'd.Range("B2").Value = 9', 'd.Rows(3).Insert: d.Range("B3").Value = 9']
S = 'ch.SeriesCollection'
READS = [f'{S}.Count', f'{S}(1).Formula', f'{S}(2).Formula', f'{S}(1).Name', f'{S}(1).Points.Count', 'ch.ChartTitle.Formula', 'ch.ChartTitle.Text', 'ot.SeriesCollection(1).Formula', 'ot.SeriesCollection(1).Points.Count', 'ch.PlotBy']


def _probe() -> str:
    lines = ['Public Function Probe(ByVal path As String, ByVal folder As String, ByVal which As Long) As String',
             '    Dim wb As Workbook, d As Worksheet, o As Worksheet, ch As Object, ot As Object, s As String, v As String', '    Application.DisplayAlerts = False', '    Set wb = Workbooks.Open(path)',
             '    Set d = wb.Worksheets("Data"): Set o = wb.Worksheets("Other"): d.Activate', '    Set ch = d.ChartObjects(1).Chart: Set ot = o.ChartObjects(1).Chart', '    On Error GoTo Failed', '    Select Case which']
    for index, edit in enumerate(EDITS):
        lines += [f'    Case {index}']
        lines += [f'        {edit}'] if edit else []
    lines += ['    End Select', '    s = "ok "', 'Answered:', '    On Error Resume Next']
    for read in READS:
        lines += ['    Err.Clear', '    v = ""', f'    v = {read}', '    If Err.Number Then v = "E" & Err.Number', '    s = s & "|" & v']
    lines += ['    On Error GoTo 0', '    Probe = s', '    wb.SaveAs folder & "\\c" & which & ".xlsx", xlOpenXMLWorkbook', '    wb.Close False', '    Application.DisplayAlerts = True', '    Exit Function',
              'Failed:', '    s = "Err " & Err.Number & " "', '    Resume Answered', 'End Function']
    return '\n'.join(lines) + '\n'


def chart_data(path: Path) -> list[object]:
    """Each chart part's references with their cached points, the first chart's first: what Excel draws from."""
    data: list[object] = []
    with zipfile.ZipFile(path) as package:
        names = sorted((name for name in package.namelist() if re.fullmatch(r'xl/charts/chart\d+\.xml', name)),
                       key=lambda name: int(re.findall(r'\d+', name)[0]))
        for name in names:
            xml = package.read(name).decode('utf-8')
            found = re.findall(r'<c:f>(.*?)</c:f>(?:<c:(?:num|str)Cache>(.*?)</c:(?:num|str)Cache>)?', xml, re.DOTALL)
            data.append(sorted([reference, re.findall(r'<c:v>(.*?)</c:v>', cache)] for reference, cache in found))
    return sorted(data, key=len, reverse=True)


def build(path: Path) -> None:
    sys.path.insert(0, str(ROOT / 'src'))
    from pyopenvba.apps.excel import ExcelApplication

    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(BUILD, name='Builder')
    app.run('Builder.Build')
    app.save(path)


def main() -> None:
    source = _probe()
    answers: list[str] = []
    with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        base = Path(folder) / 'base.xlsx'
        build(base)
        excel.new_document()
        for index in range(len(EDITS)):
            result = excel.run_vba(source, 'Probe', (str(base), folder, index), timeout=120.0)
            assert result.ok and not result.dialogs, result
            answers.append(str(result.value))
        parts = [chart_data(Path(folder) / f'c{index}.xlsx') for index in range(len(EDITS))]
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'reads': READS, 'source': source, 'build': BUILD,
              'cases': [[edit, answer, part] for edit, answer, part in zip(EDITS, answers, parts, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for edit, answer, part in zip(EDITS, answers, parts, strict=True):
        print(edit, '=>', answer)
        print('    ', str(part)[:260])


if __name__ == '__main__':
    main()
