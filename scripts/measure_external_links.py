"""Formulas that read another, closed workbook: what Range.Formula shows and what they calculate to.

Source.xlsx holds the cells and the name the formulas read; Dest.xlsx reads
them, and both are saved with both open. With Source closed, each Dest
formula is read back through Formula and FormulaR1C1 and as its cached
value; then Dest's B1 changes and the sheet recalculates from the link's
cache. The folder both files lived in is written as <folder>, so no
machine's path is kept. The parts Dest was saved with are recorded the
same way, for the model to open.

    python scripts/measure_external_links.py

writes tests/fixtures/external_links.json, which
tests/test_excel_external_links.py replays.
"""
from __future__ import annotations

import base64
import json
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/external_links.json'

#: Each formula, as written with Source open.
FORMULAS = [
    '=[Source.xlsx]Sheet1!A1', '=[Source.xlsx]Sheet1!A2*B1', '=SUM([Source.xlsx]Sheet1!A1:A5)',
    '=VLOOKUP("c",[Source.xlsx]Sheet1!B1:C5,2,FALSE)', '=INDEX([Source.xlsx]Sheet1!A1:A5,3)',
    '=Source.xlsx!Rate*100', '=[Source.xlsx]Other!A1*2', "='[Source.xlsx]Sheet1:Other'!A1",
    '=SUM([Source.xlsx]Sheet1:Other!A1)', '=[Source.xlsx]Sheet1!A6', '=[Source.xlsx]Sheet1!D1',
    '=ISREF([Source.xlsx]Sheet1!A1)', '=SUM([Source.xlsx]Sheet1!A:A)', '=COUNTBLANK([Source.xlsx]Sheet1!A1:A6)',
    '=ROW([Source.xlsx]Sheet1!A3)', '=COUNTIF([Source.xlsx]Sheet1!A1:A5,">2")', '=OFFSET([Source.xlsx]Sheet1!A1,1,0)',
    '=[Source.xlsx]Sheet1!$A$1', '=[Source.xlsx]Sheet1!A1+[Source.xlsx]Sheet1!$C$2',
]


def _literal(text: str) -> str:
    return '"' + text.replace('"', '""') + '"'


def _probe(folder: Path) -> str:
    source, dest = (str(folder / name).replace('"', '""') for name in ('Source.xlsx', 'Dest.xlsx'))
    lines = ['Function Probe() As String', 'Dim src As Workbook, dst As Workbook, ws As Worksheet, s As String, i As Long',
             'Application.DisplayAlerts = False', 'Set src = Workbooks.Add(xlWBATWorksheet)',
             'src.Worksheets(1).Name = "Sheet1"', 'src.Worksheets.Add(After:=src.Worksheets(1)).Name = "Other"',
             'With src.Worksheets("Sheet1")']
    for row, (number, letter, tens) in enumerate([(1, 'a', 10), (2, 'b', 20), (3, 'c', 30), (4, 'd', 40), (5, 'e', 50)], 1):
        lines.append(f'.Range("A{row}").Value = {number}: .Range("B{row}").Value = "{letter}": .Range("C{row}").Value = {tens}')
    lines += ['.Range("D1").Value = "x"', 'End With', 'src.Worksheets("Other").Range("A1").Value = 7',
              'src.Names.Add Name:="Rate", RefersTo:="=Sheet1!$C$1"',
              f'src.SaveAs Filename:="{source}", FileFormat:=xlOpenXMLWorkbook',
              'Set dst = Workbooks.Add(xlWBATWorksheet)', 'Set ws = dst.Worksheets(1)', 'ws.Name = "Dest"',
              'ws.Range("B1").Value = 10']
    for row, formula in enumerate(FORMULAS, 1):
        lines.append(f'ws.Range("A{row}").Formula = {_literal(formula)}')
    lines += [f'dst.SaveAs Filename:="{dest}", FileFormat:=xlOpenXMLWorkbook', 'src.Close False', 'dst.Close False',
              # Reopened alone, its links not updated, as the model opens it.
              f'Set dst = Workbooks.Open(Filename:="{dest}", UpdateLinks:=0)', 'Set ws = dst.Worksheets(1)',
              f'For i = 1 To {len(FORMULAS)}',
              's = s & "#" & ws.Cells(i, 1).Formula & "~" & ws.Cells(i, 1).FormulaR1C1 & "~" & CStr(ws.Cells(i, 1).Value)',
              'Next', 'ws.Range("B1").Value = 20', 'Application.Calculate',
              f'For i = 1 To {len(FORMULAS)}', 's = s & "~" & CStr(ws.Cells(i, 1).Value)', 'Next',
              'Application.CalculateFull',
              f'For i = 1 To {len(FORMULAS)}', 's = s & "~" & CStr(ws.Cells(i, 1).Value)', 'Next',
              # A macro writing to the link: by its path, then by its file's name alone.
              'ws.Range("C1").Formula = "=\'" & dst.Path & "\\[Source.xlsx]Sheet1\'!A3"',
              'ws.Range("C2").Formula = "=[Source.xlsx]Sheet1!A4+1"',
              's = s & "@" & ws.Range("C1").Formula & "~" & CStr(ws.Range("C1").Value) & "~" & ws.Range("C2").Formula',
              's = s & "~" & CStr(ws.Range("C2").Value)',
              # A row put in above: this workbook's references move, the link's do not.
              'ws.Rows(1).Insert',
              's = s & "@" & ws.Range("A3").Formula & "~" & ws.Range("D2").Formula & "~" & CStr(ws.Range("A3").Value)',
              'dst.Close False', 'Application.DisplayAlerts = True', 'Probe = Mid(s, 2)', 'End Function']
    return '\n'.join(lines) + '\n'


def main() -> None:
    with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=30.0)) as excel:
        excel.new_document()
        place = Path(folder)
        macro = _probe(place)
        result = excel.run_vba(macro, 'Probe', timeout=120.0, line_numbers=False)
        assert result.ok and not result.dialogs, result
        text = str(result.value)
        with zipfile.ZipFile(place / 'Dest.xlsx') as package:
            parts = {name: package.read(name) for name in package.namelist()}
        folder_text = str(place)
        record = {
            'measured_at': datetime.now(timezone.utc).isoformat(),
            'formulas': FORMULAS,
            'result': text.replace(folder_text + '\\', '<folder>\\').replace(folder_text, '<folder>'),
            # Each part as Excel saved it, its folder written <folder>; binary parts stay base64.
            'dest_parts': {name: _scrubbed(data, folder_text) for name, data in parts.items()},
        }
    OUT.write_text(json.dumps(record, indent=2) + '\n', encoding='utf-8')
    print(record['result'][:3000])


def _scrubbed(data: bytes, folder: str) -> str:
    try:
        text = data.decode('utf-8')
    except UnicodeDecodeError:
        return 'base64:' + base64.b64encode(data).decode('ascii')
    for spelling in (folder, folder.replace('\\', '/'), 'file:///' + folder.replace('\\', '/')):
        text = text.replace(spelling, '<folder>')
    return text


if __name__ == '__main__':
    main()
