"""The heights tests/fixtures/wrapped_curves_predictions.json predicted, read from live Excel after they were committed; and whether a save and reopen carries a wrapped row's height. Needs Excel at 96 DPI, checked in each call.

    python scripts/measure_wrapped_curves.py

Writes tests/fixtures/wrapped_curves.json. It uses no hyperlink, mail or send member.
"""
from __future__ import annotations

import json
import re
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

FIXTURES = Path(__file__).resolve().parents[1] / 'tests/fixtures'
LONG = 'one two three four five six seven'
SOURCE = r'''Public Function Go(ByVal text As String, ByVal font As String, ByVal size As Double, ByVal bold As Boolean, ByVal columns As String) As String
    Dim ws As Object, r As String, one As Variant
    Set ws = ActiveSheet
    ws.Cells.Clear
    With ws.Range("A1")
        .Value = "'" & text
        .WrapText = True
        .Font.Name = font
        .Font.Size = size
        .Font.Bold = bold
    End With
    For Each one In Split(columns, ",")
        ws.Columns(1).ColumnWidth = 0.1
        Do While ws.Columns(1).Width * 4 / 3 < Val(one) - 0.01
            ws.Columns(1).ColumnWidth = ws.Columns(1).ColumnWidth + 1 / 14
        Loop
        ws.Rows(1).AutoFit
        r = r & "|" & CLng(ws.Columns(1).Width * 4 / 3) & "," & ws.Rows(1).RowHeight
    Next
    Go = ActiveSheet.StandardHeight & r
End Function
'''
#: A saved book, opened: does its height stand, does it stay automatic, and what makes it work the height out again?
HISTORY = {
    'stands': [],
    'wider then saved': ['ws.Columns(1).ColumnWidth = 40'],
    'fixed': ['ws.Rows(1).RowHeight = 40'],
    'autofit after opening': ['ws.Columns(1).ColumnWidth = 40', '@ws.Rows(1).AutoFit'],
    'value after opening': ['ws.Columns(1).ColumnWidth = 40', '@ws.Range("A1").Value = "one two"'],
    'other cell after opening': ['ws.Columns(1).ColumnWidth = 40', '@ws.Range("B1").Value = "x"'],
    'font after opening': ['ws.Columns(1).ColumnWidth = 40', '@ws.Range("A1").Font.Bold = True'],
    'width after opening': ['@ws.Columns(1).ColumnWidth = 40'],
    'fixed then value after opening': ['ws.Rows(1).RowHeight = 40', '@ws.Range("A1").Value = "one two"'],
}


def history_source(name: str) -> str:
    before = [one for one in HISTORY[name] if not one.startswith('@')]
    after = [one[1:] for one in HISTORY[name] if one.startswith('@')]
    return '\n'.join(['Public Function Go(ByVal folder As String) As String', '    Dim wb As Workbook, ws As Object, r As String', '    Application.DisplayAlerts = False', '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set ws = wb.Worksheets(1)',
                      '    ws.Range("A1").WrapText = True', f'    ws.Range("A1").Value = "{LONG}"', *[f'    {one}' for one in before], '    r = ws.Rows(1).RowHeight', '    wb.SaveAs folder & "\\saved.xlsx", xlOpenXMLWorkbook', '    wb.Close False',
                      '    Set wb = Workbooks.Open(folder & "\\saved.xlsx")', '    Set ws = wb.Worksheets(1)', '    r = r & "|" & ws.Rows(1).RowHeight', *[f'    {one}' for one in after], '    r = r & "|" & ws.Rows(1).RowHeight', '    wb.Close False', '    Application.DisplayAlerts = True',
                      '    Go = ActiveSheet.StandardHeight & "|" & r', 'End Function']) + '\n'


#: The file a program writes with no height for a wrapped row: does Excel work it out on opening?
NO_HEIGHT = '\n'.join(['Public Function Go(ByVal path As String) As String', '    Dim wb As Workbook, r As String', '    Application.DisplayAlerts = False', '    Set wb = Workbooks.Open(path)', '    r = wb.Worksheets(1).Rows(1).RowHeight', '    wb.Close False', '    Application.DisplayAlerts = True', '    Go = r', 'End Function']) + '\n'


def row_markup(path: Path) -> str:
    with zipfile.ZipFile(path) as package:
        sheet = package.read('xl/worksheets/sheet1.xml').decode('utf-8')
    found = re.search(r'<row r="1"[^>]*>', sheet)
    return found.group() if found else ''


def without_height(path: Path, out: Path) -> None:
    with zipfile.ZipFile(path) as package:
        parts = {info.filename: package.read(info.filename) for info in package.infolist()}
    sheet = parts['xl/worksheets/sheet1.xml'].decode('utf-8')
    parts['xl/worksheets/sheet1.xml'] = re.sub(r'(<row r="1"[^>]*?) ht="[^"]*"', r'\1', sheet).encode('utf-8')
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as package:
        for name, data in parts.items():
            package.writestr(name, data)


def main() -> None:
    record = json.loads((FIXTURES / 'wrapped_curves_predictions.json').read_text(encoding='utf-8'))
    columns = ','.join(str(one) for one in record['columns'])
    curves = []
    history: dict[str, dict[str, str]] = {}
    with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        for name in HISTORY:
            result = excel.run_vba(history_source(name), 'Go', (folder,), timeout=60.0, line_numbers=False)
            assert result.ok and not result.dialogs, (name, result)
            height, *found = str(result.value).split('|')
            assert height == '15', f'Excel is not at 96 DPI: StandardHeight reads {height}'
            history[name] = {'set, opened, after': '|'.join(found), 'row': row_markup(Path(folder) / 'saved.xlsx')}
            if name == 'wider then saved':
                without_height(Path(folder) / 'saved.xlsx', Path(folder) / 'bare.xlsx')
                bare = excel.run_vba(NO_HEIGHT, 'Go', (str(Path(folder) / 'bare.xlsx'),), timeout=60.0, line_numbers=False)
                assert bare.ok, bare
                history['no height in the file'] = {'set, opened, after': str(bare.value), 'row': row_markup(Path(folder) / 'bare.xlsx')}
        for case in record['cases']:
            result = excel.run_vba(SOURCE, 'Go', (case['text'], case['font'], case['size'], case['bold'], columns), timeout=120.0, line_numbers=False)
            assert result.ok and not result.dialogs, (case['text'], result)
            height, *steps = str(result.value).split('|')
            assert height == '15', f'Excel is not at 96 DPI: StandardHeight reads {height}'
            curves.append({one.split(',')[0]: float(one.split(',')[1]) for one in steps})
    (FIXTURES / 'wrapped_curves.json').write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'curves': curves, 'history': history}, indent=0) + '\n', encoding='utf-8')
    good = bad = 0
    missed: dict[str, int] = {}
    for case, curve in zip(record['cases'], curves, strict=True):
        for column, tall in curve.items():
            if column in case['predicted'] and abs(case['predicted'][column] - tall) < 1e-6:
                good += 1
            else:
                bad += 1
                key = f"{case['font']} {case['size']} {case['bold']} {case['text']}"
                missed[key] = missed.get(key, 0) + 1
    print('good', good, 'off', bad)
    for key, count in sorted(missed.items(), key=lambda one: -one[1])[:40]:
        print(count, key)
    for name, found in history.items():
        print(name, found)


if __name__ == '__main__':
    main()
