"""What earns a loaded chart its automatic title, measured live on a book Excel saved, and on the same book with the chart's
autoTitleDeleted taken out as a chart not written by Excel may have it. Each case answers, and the chart parts Excel then saves:

    python scripts/measure_chart_loaded_title.py

writes tests/fixtures/chart_loaded_title.json and the books beside it, which tests/test_excel_chart_loaded_title.py replays.
"""

from __future__ import annotations

import datetime
import json
import re
import shutil
import tempfile
import zipfile
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

FIXTURES = Path(__file__).resolve().parents[1] / 'tests/fixtures'
SOURCE = r'''
Public Function Make(ByVal path As String, ByVal two As Boolean) As String
    Dim wb As Workbook, ws As Worksheet, c As Object
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Worksheets(1)
    ws.Range("A1:C1").Value = Array("", "Rev", "Cost")
    ws.Range("A2:C2").Value = Array("a", 10, 4)
    ws.Range("A3:C3").Value = Array("b", 20, 6)
    Set c = ws.ChartObjects.Add(48, 96, 240, 144).Chart
    c.SetSourceData ws.Range("A1:B3")
    If two Then ws.ChartObjects.Add(48, 300, 240, 144).Chart.SetSourceData ws.Range("A1:B3")
    wb.SaveAs path, xlOpenXMLWorkbook
    wb.Close False
End Function

Public Function Act(ByVal path As String, ByVal out As String, ByVal which As Long) As String
    Dim wb As Workbook, ws As Worksheet, c As Object, s As String
    Set wb = Workbooks.Open(path)
    Set ws = wb.Worksheets(1)
    Set c = ws.ChartObjects(1).Chart
    On Error Resume Next
    Select Case which
    Case 1: ws.ChartObjects(1).Activate
    Case 2: ws.ChartObjects(1).Select
    Case 3: c.ChartArea.Select
    Case 4: ws.ChartObjects.Select
    Case 5: ws.ChartObjects(1).Activate: c.ChartArea.Select
    Case 6: c.PlotArea.Select
    Case 7: s = c.ChartTitle.Text
    Case 8: s = c.HasTitle
    Case 9: ws.Copy After:=ws
    Case 10: c.Location 2, ws.Name
    End Select
    s = Err.Number & " " & s
    Err.Clear
    s = s & " " & TypeName(c.HasTitle) & " " & c.HasTitle & " " & Err.Number
    On Error GoTo 0
    Application.DisplayAlerts = False
    wb.SaveAs out, xlOpenXMLWorkbook
    Application.DisplayAlerts = True
    wb.Close False
    Act = s
End Function
'''
LABELS = ('nothing', 'ChartObject.Activate', 'ChartObject.Select', 'ChartArea.Select', 'ChartObjects.Select', 'Activate then ChartArea.Select',
          'PlotArea.Select', 'ChartTitle.Text', 'read HasTitle', 'Worksheet.Copy', 'Location on its own sheet')
ABSOLUTE = re.compile(r'<mc:AlternateContent[^>]*><mc:Choice Requires="x15"><x15ac:absPath.*?</mc:AlternateContent>', re.S)


def rewritten(source: Path, out: Path, change: dict[str, object]) -> None:
    """The book with its folder taken out, and the chart's flag too when asked."""
    with zipfile.ZipFile(source) as package, zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as target:
        for item in package.infolist():
            data = package.read(item.filename)
            if item.filename == 'xl/workbook.xml':
                data = ABSOLUTE.sub('', data.decode('utf-8')).encode('utf-8')
            elif change.get('no_flag') and re.fullmatch(r'xl/charts/chart\d+\.xml', item.filename):
                data = re.sub(r'<c:autoTitleDeleted [^>]*/>', '', data.decode('utf-8')).encode('utf-8')
            target.writestr(item, data)


def charts(path: Path) -> dict[str, str]:
    with zipfile.ZipFile(path) as package:
        return {name: package.read(name).decode('utf-8') for name in package.namelist() if re.fullmatch(r'xl/charts/chart\d+\.xml', name)}


def main() -> None:
    folder = Path(tempfile.mkdtemp())
    books = {'chart_loaded_title_one.xlsx': (False, {}), 'chart_loaded_title_two.xlsx': (True, {}),
             'chart_loaded_title_no_flag.xlsx': (False, {'no_flag': True})}
    record: dict[str, object] = {'measured_at': datetime.datetime.now(datetime.timezone.utc).isoformat(), 'source': SOURCE, 'labels': LABELS}
    cases: list[list[object]] = []
    with ExcelSession(HarnessConfig(lock_wait_s=600)) as excel:
        for name, (two, change) in books.items():
            made = folder / f'made{name}'
            done = excel.run_vba(SOURCE, 'Make', (str(made), two), timeout=120.0, line_numbers=False)
            assert done.outcome == 'passed', done.message
            rewritten(made, FIXTURES / name, change)
            for which, label in enumerate(LABELS):
                copy, out = folder / 'in.xlsx', folder / f'{name}{which}.xlsx'
                shutil.copy(FIXTURES / name, copy)
                done = excel.run_vba(SOURCE, 'Act', (str(copy), str(out), which), timeout=120.0, line_numbers=False)
                assert done.outcome == 'passed', (label, done.message)
                cases.append([name, which, done.value, charts(out)])
                print(name, label, '=>', done.value, ['title' if '<c:chart><c:title>' in text else 'none' for text in charts(out).values()])
    record['cases'] = cases
    (FIXTURES / 'chart_loaded_title.json').write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8', newline='\n')


if __name__ == '__main__':
    main()
