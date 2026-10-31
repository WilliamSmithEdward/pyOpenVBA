"""Measure Workbook.PrecisionAsDisplayed: what turning it on does to the constants and formula results of cells
in each kind of number format, what later entries and format changes do, turning it off again, and what a save
writes.

Each case runs on a new workbook; cell A1 holds a value in a format, and the answer is A1's Value2 bits and Text
after the statements, with B1's where a formula reads A1. Some cases save the workbook and workbook.xml's calcPr is
recorded.

    python scripts/measure_precision_as_displayed.py

writes tests/fixtures/precision_as_displayed.json, which tests/test_excel_precision_as_displayed.py replays.
"""
from __future__ import annotations

import json
import re
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/precision_as_displayed.json'
#: (value as VBA, number format) put in A1 before PrecisionAsDisplayed is turned on.
FORMATS: list[tuple[str, str]] = [
    ('1.23456', '0.00'), ('1.235', '0.00'), ('-1.235', '0.00'), ('2.5', '0'), ('3.5', '0'), ('1234.5678', '#,##0'),
    ('1234567.89', '#,##0,'), ('1234567.89', '0.0,,'), ('0.123456', '0%'), ('0.123456', '0.0%'),
    ('123456.789', '0.00E+00'), ('0.333333333333333', '# ?/?'), ('1.33333333333333', '# ??/??'),
    ('1.23456', 'General'), ('1.23456789012345', 'General'), ('123456789012.345', 'General'), ('1.23456', '@'),
    ('45293.75', 'm/d/yyyy'), ('45293.7512345', 'h:mm'), ('45293.7512345', 'h:mm:ss'),
    ('45293.7512345', 'm/d/yyyy h:mm'), ('0.75123456', '[h]:mm'), ('1.23456', '$#,##0.00'),
    ('-1.23456', '0.00;(0.00)'), ('-1.23456', '0.00;-0'), ('0.004', '0.00'), ('-0.004', '0.00'),
    ('1.23456', '0.00" units"'), ('1.23456', '0.0#'), ('1.2', '0.0#'), ('12.3456', '?0.0?'),
    ('1.23456', '0.00;;;'), ('1.23456', ';;;'), ('1.23456', '[Red]0.00'), ('1.23456', '[<2]0.0;0.00'),
    ('1E+20', '0.00'), ('1.23456', '0.00_)'), ('1.23456', '0.00*-'), ('true', '0.00'), ('"text"', '0.00'),
    ('1.23456', '0.000000000000000000'),
]
#: Each case: statements on wb, ws after the setup; the cells answered.
CASES: list[tuple[str, str, list[str], str]] = []
for value, number_format in FORMATS:
    CASES.append((value, number_format, ['wb.PrecisionAsDisplayed = True'], 'A1'))
CASES += [
    ('1.23456', '0.00', [], 'A1'),
    ('1.23456', '0.00', ['wb.PrecisionAsDisplayed = True', 'wb.PrecisionAsDisplayed = False'], 'A1'),
    ('1.23456', '0.00', ['ws.Range("B1").Formula = "=A1*3"', 'wb.PrecisionAsDisplayed = True'], 'A1:B1'),
    ('1.23456', 'General', ['ws.Range("B1").Formula = "=A1*3"', 'ws.Range("B1").NumberFormat = "0.0"',
                            'wb.PrecisionAsDisplayed = True'], 'A1:B1'),
    ('1.23456', 'General', ['ws.Range("B1").Formula = "=1/3"', 'ws.Range("B1").NumberFormat = "0.00"',
                            'ws.Range("C1").Formula = "=B1*3"', 'wb.PrecisionAsDisplayed = True'], 'A1:C1'),
    ('1.23456', 'General', ['wb.PrecisionAsDisplayed = True', 'ws.Range("B1").NumberFormat = "0.00"',
                            'ws.Range("B1").Value = 2.555'], 'A1:B1'),
    ('1.23456', 'General', ['wb.PrecisionAsDisplayed = True', 'ws.Range("A1").NumberFormat = "0.0"'], 'A1'),
    ('1.23456', 'General', ['wb.PrecisionAsDisplayed = True', 'ws.Range("A1").NumberFormat = "0.0"',
                            'ws.Calculate'], 'A1'),
    ('1.23456', 'General', ['wb.PrecisionAsDisplayed = True', 'ws.Range("B1").NumberFormat = "0.00"',
                            'ws.Range("B1").Formula = "=1/3"', 'ws.Range("C1").Formula = "=B1*3"'], 'A1:C1'),
    ('1.23456', 'General', ['wb.PrecisionAsDisplayed = True', 'ws.Range("B1").NumberFormat = "0.00"',
                            'ws.Range("B1").Formula = "=A1"', 'ws.Range("A1").Value = 9.8765'], 'A1:B1'),
    ('1.23456', '0.00', ['wb.PrecisionAsDisplayed = True', 'ws.Range("B1").Formula = "=A1*3"'], 'A1:B1'),
    ('1.23456', '0.00', ['wb.PrecisionAsDisplayed = True', 'ws.Range("B1").Value = "=1/3"'], 'A1:B1'),
    ('1.23456', '0.00', ['wb.PrecisionAsDisplayed = True', 'ws.Range("A1").Value = 7.777'], 'A1'),
    ('1.23456', '0.00', ['wb.PrecisionAsDisplayed = True', 'ws.Range("A1").Value2 = 7.777'], 'A1'),
    ('1.23456', '0.00', ['wb.PrecisionAsDisplayed = True', 'ws.Range("A1").Formula = "7.777"'], 'A1'),
    ('1.23456', '0.00', ['wb.PrecisionAsDisplayed = True', 'ws.Range("A1:A2").Value = 7.777'], 'A1:A2'),
    ('1.23456', '0.00', ['ws.Range("A2").Value = 9.87654', 'ws.Range("A2").NumberFormat = "0.0"',
                         'ws.Range("A3").Value = 5.55555', 'Set ws2 = wb.Worksheets.Add',
                         'ws2.Range("A1").Value = 3.33333', 'ws2.Range("A1").NumberFormat = "0.0"',
                         'wb.PrecisionAsDisplayed = True', 'ws.Range("A3").Value = ws.Range("A3").Value * 1',
                         'Set c = ws2.Range("A1")'], 'A1:A3'),
    ('1.23456', '0.00', ['wb.PrecisionAsDisplayed = True', 'Set c = wb.Worksheets.Add.Range("A1")',
                         'c.NumberFormat = "0.0"', 'c.Value = 3.33333'], 'A1'),
    ('1.23456', '0.00', ['wb.PrecisionAsDisplayed = 1'], 'A1'),
    ('1.23456', '0.00', ['wb.PrecisionAsDisplayed = "True"'], 'A1'),
    ('1.23456', '0.00', ['wb.PrecisionAsDisplayed = True', 'r = wb.PrecisionAsDisplayed'], 'A1'),
    ('1.23456', '0.00', ['wb.PrecisionAsDisplayed = True', 'wb.SaveAs folder & "\\on.xlsx"'], 'A1'),
    ('1.23456', '0.00', ['wb.SaveAs folder & "\\off.xlsx"'], 'A1'),
    ('1.23456', '0.00', ['wb.PrecisionAsDisplayed = True', 'wb.SaveAs folder & "\\reopen.xlsx"', 'wb.Close',
                         'Set wb = Workbooks.Open(folder & "\\reopen.xlsx")', 'Set ws = wb.Worksheets(1)',
                         'r = wb.PrecisionAsDisplayed', 'ws.Range("B1").NumberFormat = "0.0"',
                         'ws.Range("B1").Value = 2.2222'], 'A1:B1'),
    ('1.23456', '0.00', ['ws.Range("B1").Formula = "=SUM(A1,0.001)"', 'ws.Range("B1").NumberFormat = "0.00"',
                         'Application.Calculation = xlCalculationManual', 'wb.PrecisionAsDisplayed = True',
                         'r = ws.Range("B1").Value2', 'Application.Calculation = xlCalculationAutomatic'], 'A1:B1'),
    ('1.23456', '0.00', ['ws.Range("A2:A3").Value = 4.4444', 'ws.Range("A2:A3").NumberFormat = "0.0"',
                         'ws.Range("B1").FormulaArray = "=A1:A3*2"', 'ws.Range("B1:B3").NumberFormat = "0"',
                         'wb.PrecisionAsDisplayed = True'], 'A1:B1'),
]


def _probe() -> str:
    lines = ['Private Function Bits(ByVal v As Variant) As String',
             '    If VarType(v) = vbDouble Then',
             '        Dim d As Double: d = v',
             '        Bits = "D:" & Format(d, "0.000000000000000E+00")', '    Else',
             '        Bits = TypeName(v) & ":" & CStr(v)', '    End If', 'End Function', '',
             'Private Function Tried(ByVal which As Long, ByVal folder As String) As String',
             '    Dim wb As Workbook, ws As Worksheet, ws2 As Worksheet, c As Range, cell As Range, s As String',
             '    Dim r As Variant', '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set ws = wb.Worksheets(1)',
             '    r = Empty', '    On Error GoTo Failed', '    Select Case which']
    for index, (value, number_format, statements, _read) in enumerate(CASES):
        lines.append(f'    Case {index}')
        lines.append(f'        ws.Range("A1").NumberFormat = "{number_format.replace(chr(34), chr(34) * 2)}"')
        lines.append(f'        ws.Range("A1").Value = {value}')
        lines += [f'        {statement}' for statement in statements]
    lines += ['    End Select']
    lines += ['    Select Case which']
    for index, (_value, _format, _statements, read) in enumerate(CASES):
        lines.append(f'    Case {index}')
        lines.append(f'        For Each cell In ws.Range("{read}")')
        lines.append('            s = s & ";" & cell.Address(False, False) & "=" & Bits(cell.Value2) & " [" & '
                     'cell.Text & "]"')
        lines.append('        Next')
    lines += ['    End Select',
              '    If Not c Is Nothing Then s = s & ";c=" & Bits(c.Value2) & " [" & c.Text & "]"',
              '    If Not IsEmpty(r) Then s = s & ";r=" & Bits(r)',
              '    Tried = "P=" & wb.PrecisionAsDisplayed & s', 'Done:', '    wb.Close False', '    Exit Function',
              'Failed:', '    Tried = "Err " & Err.Number & " " & Err.Description', '    Resume Done', 'End Function',
              '',
              'Public Function Probe(ByVal folder As String, ByVal first As Long, ByVal last As Long) As String',
              '    Dim s As String, i As Long', '    Application.DisplayAlerts = False', '    For i = first To last',
              '        s = s & "@@@" & Tried(i, folder)', '    Next', '    Application.DisplayAlerts = True',
              '    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def _calc(path: Path) -> str:
    with zipfile.ZipFile(path) as package:
        workbook = package.read('xl/workbook.xml').decode('utf-8')
    return (re.search(r'<calcPr\b[^>]*/?>', workbook) or [''])[0]


def main() -> None:
    source = _probe()
    with tempfile.TemporaryDirectory() as folder:
        with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
            excel.new_document()
            result = excel.run_vba(source, 'Probe', (folder, 0, len(CASES) - 1), timeout=300.0)
        assert result.ok and not result.dialogs, result
        answers = str(result.value).split('@@@')
        calc = {name: _calc(Path(folder) / name) for name in ('on.xlsx', 'off.xlsx', 'reopen.xlsx')
                if (Path(folder) / name).exists()}
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source,
              'cases': [[value, number_format, statements, read, answer]
                        for (value, number_format, statements, read), answer in zip(CASES, answers, strict=True)],
              'calc': calc}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for index, (value, number_format, statements, _read, answer) in enumerate(record['cases']):
        print(index, value, number_format, '|', (statements or [''])[-1][:50], '->', answer)
    print(calc)


if __name__ == '__main__':
    main()
