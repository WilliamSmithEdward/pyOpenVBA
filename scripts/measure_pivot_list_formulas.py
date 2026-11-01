"""PivotTable.ListFormulas: the sheet it makes, where that sheet goes and what it is named, which sheet is active
after, and every cell it fills with its formula, number format and boldness; and the errors. Column widths are left
out: Excel fits them to its text at this machine's DPI.

Each case makes a workbook with a list of regions, products and amounts on its first sheet, builds a pivot table on a
second sheet with Region as a row field and the sum of Amount as data, runs the case's statements with pt as the
table, then reads every sheet's name in order, the active sheet's name, and, when the active sheet is no longer the
report, each filled cell of it.

    python scripts/measure_pivot_list_formulas.py

writes tests/fixtures/pivot_list_formulas.json, which tests/test_excel_pivot_list_formulas.py replays.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/pivot_list_formulas.json'
BONUS = 'pt.CalculatedFields.Add "Bonus", "=Amount*0.1"'
CASES = [
    'pt.ListFormulas',
    f'{BONUS}: pt.ListFormulas',
    f'{BONUS}: pt.CalculatedFields.Add "Twice", "=Amount*2": pt.ListFormulas',
    f'{BONUS}: pt.ListFormulas: pt.ListFormulas',
    'pt.PivotFields("Region").CalculatedItems.Add "Both", "=East+West": pt.ListFormulas',
    f'{BONUS}: pt.PivotFields("Region").CalculatedItems.Add "Both", "=East+West": pt.ListFormulas',
    f'{BONUS}: s = TypeName(pt.ListFormulas)',
    f'{BONUS}: wb.Sheets("data").Activate: pt.ListFormulas',
]


def _probe() -> str:
    lines = ['Private Function Tried(ByVal which As Long) As String',
             '    Dim wb As Workbook, ws As Worksheet, pt As Object, c As Range, s As String, sh As Object',
             '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set ws = wb.Worksheets(1)', '    ws.Name = "data"',
             '    ws.Range("A1:C1").Value = Array("Region", "Product", "Amount")',
             '    ws.Range("A2:C2").Value = Array("East", "Pen", 10)',
             '    ws.Range("A3:C3").Value = Array("West", "Ink", 20)',
             '    Set ws = wb.Worksheets.Add(After:=ws)', '    ws.Name = "report"',
             '    Set pt = wb.PivotCaches.Create(xlDatabase, "data!R1C1:R3C3").CreatePivotTable(ws.Range("A3"), "P")',
             '    pt.PivotFields("Region").Orientation = xlRowField',
             '    pt.AddDataField pt.PivotFields("Amount"), "Sum of Amount", xlSum',
             '    On Error GoTo Failed', '    Select Case which']
    for index, statements in enumerate(CASES):
        lines += [f'    Case {index}', f'        {statements}']
    lines += ['    End Select',
              '    For Each sh In wb.Sheets: s = s & "|" & sh.Name: Next', '    s = s & "|active=" & ActiveSheet.Name',
              '    If ActiveSheet.Name <> "report" Then',
              '        For Each c In ActiveSheet.UsedRange.Cells',
              '            If Not IsEmpty(c.Value) Then s = s & "|" & c.Address(False, False) & ":" & c.Formula & ":" & '
              'c.NumberFormat & ":" & c.Font.Bold',
              '        Next', '    End If',
              '    Tried = "ok " & s', 'Done:', '    wb.Close False', '    Exit Function', 'Failed:',
              '    Tried = "Err " & Err.Number', '    Resume Done', 'End Function', '',
              'Public Function Probe() As String', '    Dim i As Long, s As String',
              f'    For i = 0 To {len(CASES) - 1}', '        s = s & "@@@" & Tried(i)', '    Next',
              '    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def main() -> None:
    source = _probe()
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(source, 'Probe', timeout=280.0)
        assert result.ok and not result.dialogs, result
        answers = str(result.value).split('@@@')
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source,
              'cases': [list(pair) for pair in zip(CASES, answers, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for case, answer in record['cases']:
        print(case, '->', answer)


if __name__ == '__main__':
    main()
