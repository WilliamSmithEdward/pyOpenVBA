"""Measure Range.Errors: which of Excel's error checks flag a cell, what Ignore does to them, and what a save
writes for an ignored check.

Each case runs on a new workbook and answers each of the ten Errors(i).Value of the cell it names as 1 or 0, then
Ignore; some save the workbook and the sheet part's ignoredErrors is recorded. Application.ErrorCheckingOptions is
only read: its settings persist in the registry.

    python scripts/measure_range_errors.py

writes tests/fixtures/range_errors.json, which tests/test_excel_range_errors.py replays.
"""
from __future__ import annotations

import json
import re
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/range_errors.json'
OPTIONS = ['BackgroundChecking', 'EvaluateToError', 'TextDate', 'NumberAsText', 'InconsistentFormula',
           'OmittedCells', 'UnlockedFormulaCells', 'EmptyCellReferences', 'ListDataValidation',
           'InconsistentTableFormula', 'MisleadingNumberFormats']
#: Each case: statements on a new workbook wb with sheet ws; the cell whose checks are answered; the file saved.
CASES: list[tuple[list[str], str, str]] = [
    (['ws.Range("A1").Value = 5'], 'A1', ''),
    ([], 'A1', ''),
    (['ws.Range("A1").Value = "\'123"'], 'A1', ''),
    (['ws.Range("A1").NumberFormat = "@"', 'ws.Range("A1").Value = "123"'], 'A1', ''),
    (['ws.Range("A1").Value = "\'1.5"'], 'A1', ''),
    (['ws.Range("A1").Value = "\'-7"'], 'A1', ''),
    (['ws.Range("A1").Value = "\'1E3"'], 'A1', ''),
    (['ws.Range("A1").Value = "\'$5"'], 'A1', ''),
    (['ws.Range("A1").Value = "\'12%"'], 'A1', ''),
    (['ws.Range("A1").Value = "\' 123"'], 'A1', ''),
    (['ws.Range("A1").Value = "\'1,234"'], 'A1', ''),
    (['ws.Range("A1").Value = "\'abc"'], 'A1', ''),
    (['ws.Range("A1").Value = "\'TRUE"'], 'A1', ''),
    (['ws.Range("A1").Value = "\'1/2/24"'], 'A1', ''),
    (['ws.Range("A1").Value = "\'1/2/2024"'], 'A1', ''),
    (['ws.Range("A1").Value = "\'Jan-24"'], 'A1', ''),
    (['ws.Range("A1").Value = "\'2-Jan"'], 'A1', ''),
    (['ws.Range("A1").Formula = "=1/0"'], 'A1', ''),
    (['ws.Range("A1").Formula = "=NA()"'], 'A1', ''),
    (['ws.Range("A1").Value = CVErr(xlErrNA)'], 'A1', ''),
    (['ws.Range("A1").Formula = "=B1"'], 'A1', ''),
    (['ws.Range("B1").Value = 2', 'ws.Range("A1").Formula = "=B1"'], 'A1', ''),
    (['ws.Range("A1").Formula = "=SUM(B1:B3)"'], 'A1', ''),
    (['ws.Range("A1").Formula = "=1+1"', 'ws.Range("A1").Locked = False'], 'A1', ''),
    (['ws.Range("A1").Value = 5', 'ws.Range("A1").Locked = False'], 'A1', ''),
    (['ws.Range("A1:A3").Value = Application.Transpose(Array(1, 2, 3))', 'ws.Range("B1").Formula = "=A1*2"',
      'ws.Range("B2").Formula = "=A2*3"', 'ws.Range("B3").Formula = "=A3*2"'], 'B2', ''),
    (['ws.Range("A1:A3").Value = Application.Transpose(Array(1, 2, 3))', 'ws.Range("B1").Formula = "=A1*2"',
      'ws.Range("B2").Formula = "=A2*3"', 'ws.Range("B3").Formula = "=A3*2"'], 'B1', ''),
    (['ws.Range("A1:A3").Value = Application.Transpose(Array(1, 2, 3))', 'ws.Range("A4").Formula = "=SUM(A1:A2)"'],
     'A4', ''),
    (['ws.Range("A1:A3").Value = Application.Transpose(Array(1, 2, 3))', 'ws.Range("A4").Formula = "=SUM(A1:A3)"'],
     'A4', ''),
    (['ws.Range("A1").Value = "\'123"', 'ws.Range("A1").Errors(xlNumberAsText).Ignore = True'], 'A1', 'ignored.xlsx'),
    (['ws.Range("A1:B2").Value = "\'123"', 'ws.Range("A1:B2").Errors(xlNumberAsText).Ignore = True'], 'B2',
     'ignored_area.xlsx'),
    (['ws.Range("A1").Formula = "=1/0"', 'ws.Range("A1").Errors(xlEvaluateToError).Ignore = True'], 'A1',
     'ignored_error.xlsx'),
    (['ws.Range("A1").Formula = "=1/0"', 'ws.Range("A1").Errors(1).Ignore = True',
      'ws.Range("A1").Errors(7).Ignore = True'], 'A1', 'ignored_two.xlsx'),
    (['ws.Range("A1").Formula = "=1/0"', 'ws.Range("A1").Errors(1).Ignore = True',
      'ws.Range("A1").Errors(1).Ignore = False'], 'A1', 'ignored_back.xlsx'),
    (['ws.Range("A1").Value = 5', 'ws.Range("A1").Errors(3).Ignore = True'], 'A1', 'ignored_number.xlsx'),
    (['ws.Range("A1").Value = "\'123"', 'ws.Range("C3").Errors(3).Ignore = True'], 'A1', 'ignored_elsewhere.xlsx'),
    (['ws.Range("A1").Value = "\'123"', 'ws.Range("A1").Errors(3).Ignore = True', 'ws.Range("A2").Errors(3).Ignore = True'],
     'A1', 'ignored_neighbours.xlsx'),
    (['ws.Range("A1").Value = "\'123"', 'ws.Range("A1,C1").Errors(3).Ignore = True'], 'A1', 'ignored_union.xlsx'),
    (['ws.Range("A1").Value = "\'123"', 'ws.Range("A1").Errors(3).Ignore = True', 'ws.Range("A1").Value = "\'456"'],
     'A1', ''),
    (['ws.Range("A1").Value = "\'123"', 'ws.Range("A1").Errors(3).Ignore = True', 'ws.Range("A1").ClearContents'],
     'A1', ''),
    (['ws.Range("A1").Value = "\'123"', 'ws.Range("A1").Errors(3).Ignore = True', 'ws.Range("A1").Clear'], 'A1', ''),
    (['ws.Range("A1").Value = "\'123"', 'r = ws.Range("A1:B2").Errors(3).Value'], 'A1', ''),
    (['ws.Range("B1").Value = "\'123"', 'r = ws.Range("A1:B2").Errors(3).Value'], 'A1', ''),
    (['r = ws.Range("A1").Errors(0).Value'], 'A1', ''),
    (['r = ws.Range("A1").Errors(11).Value'], 'A1', ''),
    (['r = TypeName(ws.Range("A1").Errors) & TypeName(ws.Range("A1").Errors(1)) & '
      'TypeName(ws.Range("A1").Errors(1).Parent) & TypeName(ws.Range("A1").Errors.Parent)'], 'A1', ''),
    (['r = ws.Range("A1").Errors.Item(3).Value'], 'A1', ''),
    (['ws.Range("A1").Value = "\'123"', 'ws.Range("A1").Errors(3).Ignore = "True"'], 'A1', ''),
    (['ws.Range("A1").Value = "\'123"', 'ws.Range("A1").Errors(3).Ignore = 1'], 'A1', ''),
    (['ws.Range("A1").Formula = "=""123"""'], 'A1', ''),
    (['ws.Range("A1").Formula = "=A2+1"', 'ws.Range("A2").Value = "\'5"'], 'A1', ''),
    (['ws.Range("A1").Value = "\'123"', 'ws.Protect'], 'A1', ''),
    (['ws.Range("A1").Value = "\'00123"'], 'A1', ''),
    (['ws.Range("A1").Formula = "=SQRT(-1)"'], 'A1', ''),
    (['ws.Range("A1").Formula = "=1+""a"""'], 'A1', ''),
    (['ws.Range("A1").Formula = "=nosuchname"'], 'A1', ''),
    (['ws.Range("A1").Formula = "=#REF!"'], 'A1', ''),
    (['ws.Range("A1").Formula = "=IFERROR(1/0,1)"'], 'A1', ''),
    (['ws.Range("A1").Formula = "=B1:B2 C3:C4"'], 'A1', ''),
    (['ws.Range("B1").Value = 2', 'ws.Range("A1").Formula = "=SUM(B1:B3)"'], 'A1', ''),
    (['ws.Range("B1").Value = "x"', 'ws.Range("A1").Formula = "=B1"'], 'A1', ''),
    (['ws.Range("A1").Formula = "=Sheet2!B1"'], 'A1', ''),
    (['ws.Range("A1").Formula = "=ROW(B1)"'], 'A1', ''),
    (['ws.Range("A1").Formula = "=ISBLANK(B1)"'], 'A1', ''),
    (['ws.Range("A1:A3").Value = Application.Transpose(Array(1, 2, 3))', 'ws.Range("B1").Formula = "=A1*2"',
      'ws.Range("B2").Formula = "=A2*2"', 'ws.Range("B3").Formula = "=A3*3"'], 'B3', ''),
    (['ws.Range("A1:A4").Value = Application.Transpose(Array(1, 2, 3, 4))', 'ws.Range("B1").Formula = "=A1*2"',
      'ws.Range("B2").Formula = "=A2*2"', 'ws.Range("B3").Formula = "=A3*3"'], 'B3', ''),
    (['ws.Range("A1:A3").Value = Application.Transpose(Array(1, 2, 3))', 'ws.Range("B1").Formula = "=A1*3"',
      'ws.Range("B2").Formula = "=A2*2"', 'ws.Range("B3").Formula = "=A3*2"'], 'B1', ''),
    (['ws.Range("A1:C1").Value = Array(1, 2, 3)', 'ws.Range("A2").Formula = "=A1*2"',
      'ws.Range("B2").Formula = "=B1*3"', 'ws.Range("C2").Formula = "=C1*2"'], 'B2', ''),
    (['ws.Range("A1:A3").Value = Application.Transpose(Array(1, 2, 3))', 'ws.Range("B1").Formula = "=A1*2"',
      'ws.Range("B2").Value = 5', 'ws.Range("B3").Formula = "=A3*2"'], 'B2', ''),
    (['ws.Range("A1:A3").Value = Application.Transpose(Array(1, 2, 3))', 'ws.Range("B1").Formula = "=A1*2"',
      'ws.Range("B2").Formula = "=A2*3"', 'ws.Range("B3").Formula = "=A3*4"'], 'B2', ''),
    (['ws.Range("A1:A4").Value = Application.Transpose(Array(1, 2, 3, 4))', 'ws.Range("A5").Formula = "=SUM(A2:A4)"'],
     'A5', ''),
    (['ws.Range("A1:A3").Value = Application.Transpose(Array(1, 2, 3))', 'ws.Range("A4").Formula = "=A1+A2"'],
     'A4', ''),
    (['ws.Range("A1:C1").Value = Array(1, 2, 3)', 'ws.Range("D1").Formula = "=SUM(A1:B1)"'], 'D1', ''),
    (['ws.Range("A1:A3").Value = Application.Transpose(Array(1, 2, 3))', 'ws.Range("C1").Formula = "=SUM(A1:A2)"'],
     'C1', ''),
    (['ws.Range("A1:A2").Value = Application.Transpose(Array(1, 2))', 'ws.Range("A3").Value = "x"',
      'ws.Range("A4").Formula = "=SUM(A1:A2)"'], 'A4', ''),
    (['ws.Range("A1").Value = 0.5', 'ws.Range("A1").NumberFormat = "0%"', 'ws.Range("B1").Formula = "=A1"',
      'ws.Range("B1").NumberFormat = "General"'], 'B1', ''),
    (['ws.Range("A1").Value = 0.5', 'ws.Range("A1").NumberFormat = "0%"', 'ws.Range("B1").Formula = "=A1*2"',
      'ws.Range("B1").NumberFormat = "0.00"'], 'B1', ''),
]


def _probe() -> str:
    options = ' & '.join(f'IIf(Application.ErrorCheckingOptions.{name}, "1", "0")' for name in OPTIONS)
    lines = ['Private Function Tried(ByVal which As Long, ByVal folder As String) As String',
             '    Dim wb As Workbook, ws As Worksheet, c As Range, i As Long, s As String, r As Variant',
             '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set ws = wb.Worksheets(1)', '    r = Empty',
             '    On Error GoTo Failed', '    Select Case which']
    for index, (statements, cell, saved) in enumerate(CASES):
        lines.append(f'    Case {index}')
        lines += [f'        {statement}' for statement in statements]
        lines.append(f'        Set c = ws.Range("{cell}")')
        if saved:
            lines.append(f'        wb.SaveAs folder & "\\{saved}", xlOpenXMLWorkbook')
    lines += ['    End Select', '    For i = 1 To 10', '        s = s & IIf(c.Errors(i).Value, "1", "0")', '    Next',
              '    s = s & " ignore="', '    For i = 1 To 10', '        s = s & IIf(c.Errors(i).Ignore, "1", "0")',
              '    Next', '    If Not IsEmpty(r) Then s = s & " r=" & CStr(r)', '    Tried = s', 'Done:',
              '    wb.Close False', '    Exit Function', 'Failed:',
              '    Tried = "Err " & Err.Number & " " & Err.Description', '    Resume Done', 'End Function', '',
              'Public Function Options() As String', f'    Options = {options}', 'End Function', '',
              'Public Function Probe(ByVal folder As String, ByVal first As Long, ByVal last As Long) As String',
              '    Dim s As String, i As Long', '    Application.DisplayAlerts = False', '    For i = first To last',
              '        s = s & "@@@" & Tried(i, folder)', '    Next', '    Application.DisplayAlerts = True',
              '    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def _ignored(path: Path) -> str:
    with zipfile.ZipFile(path) as package:
        xml = package.read('xl/worksheets/sheet1.xml').decode('utf-8')
    return (re.search(r'<ignoredErrors>.*?</ignoredErrors>', xml, re.DOTALL) or [''])[0]


def main() -> None:
    source = _probe()
    with tempfile.TemporaryDirectory() as folder:
        with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
            excel.new_document()
            options = excel.run_vba(source, 'Options', (), timeout=60.0)
            result = excel.run_vba(source, 'Probe', (folder, 0, len(CASES) - 1), timeout=300.0)
        assert options.ok and result.ok and not result.dialogs, (options, result)
        answers = str(result.value).split('@@@')
        ignored = {saved: _ignored(Path(folder) / saved) for _, _, saved in CASES
                   if saved and (Path(folder) / saved).exists()}
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source,
              'options': dict(zip(OPTIONS, str(options.value), strict=True)),
              'cases': [[statements, cell, saved, answer]
                        for (statements, cell, saved), answer in zip(CASES, answers, strict=True)],
              'ignored': ignored}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    print(record['options'])
    for index, (statements, cell, saved, answer) in enumerate(record['cases']):
        print(index, (statements or [''])[-1][:62], cell, '->', answer, '|', ignored.get(saved, ''))


if __name__ == '__main__':
    main()
