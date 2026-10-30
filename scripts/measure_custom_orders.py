"""Custom sort orders: Sort.SortFields.Add's CustomOrder, as text and as a list's number, and the sortState a save
writes for it; Range.Sort's OrderCustom with MatchCase, and whether it leaves a sortState; AddCustomList of a range
mixing text and numbers.

The lists are Excel's own, kept for the user, so the probe deletes every list it added before it returns.

    python scripts/measure_custom_orders.py

writes tests/fixtures/custom_orders.json, which tests/test_excel_custom_orders.py replays.
"""
from __future__ import annotations

import json
import re
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/custom_orders.json'
VALUES = 'Application.Transpose(Array("High", "low", "Zeta", "Medium", 5, "high", "Apple"))'
#: Each case: statements on a new workbook's sheet ws, run after the list Low, Medium, High is added as list n + 1;
#: the cells read back; the file saved, if any.
CASES = [
    (['ws.Range("A1:A7").Value = ' + VALUES, 'ws.Sort.SortFields.Clear',
      'ws.Sort.SortFields.Add Key:=ws.Range("A1:A7"), Order:=xlAscending, CustomOrder:="High,Medium,Low"',
      'ws.Sort.SetRange ws.Range("A1:A7")', 'ws.Sort.Header = xlNo', 'ws.Sort.Apply'], 'A1:A7', 'text.xlsx'),
    (['ws.Range("A1:A7").Value = ' + VALUES, 'ws.Sort.SortFields.Clear',
      'ws.Sort.SortFields.Add Key:=ws.Range("A1:A7"), Order:=xlDescending, CustomOrder:="High,Medium,Low"',
      'ws.Sort.SetRange ws.Range("A1:A7")', 'ws.Sort.Header = xlNo', 'ws.Sort.Apply'], 'A1:A7', 'down.xlsx'),
    (['ws.Range("A1:A7").Value = ' + VALUES, 'ws.Sort.SortFields.Clear',
      'ws.Sort.SortFields.Add Key:=ws.Range("A1:A7"), Order:=xlAscending, CustomOrder:="High, Medium,Low"',
      'ws.Sort.SetRange ws.Range("A1:A7")', 'ws.Sort.Header = xlNo', 'ws.Sort.Apply'], 'A1:A7', 'spaced.xlsx'),
    (['ws.Range("A1:A7").Value = ' + VALUES, 'ws.Sort.SortFields.Clear',
      'ws.Sort.SortFields.Add Key:=ws.Range("A1:A7"), Order:=xlAscending, CustomOrder:=n + 1',
      'ws.Sort.SetRange ws.Range("A1:A7")', 'ws.Sort.Header = xlNo', 'ws.Sort.Apply'], 'A1:A7', 'number.xlsx'),
    (['ws.Range("A1:A5").Value = Application.Transpose(Array("Wed", "Sunday", "Mon", "sat", "Tue"))',
      'ws.Sort.SortFields.Clear',
      'ws.Sort.SortFields.Add Key:=ws.Range("A1:A5"), Order:=xlAscending, CustomOrder:=1',
      'ws.Sort.SetRange ws.Range("A1:A5")', 'ws.Sort.Header = xlNo', 'ws.Sort.Apply'], 'A1:A5', 'days.xlsx'),
    (['ws.Range("A1:A7").Value = ' + VALUES, 'ws.Sort.SortFields.Clear',
      'ws.Sort.SortFields.Add Key:=ws.Range("A1:A7"), Order:=xlAscending, CustomOrder:="High,Medium,Low"',
      'ws.Sort.SetRange ws.Range("A1:A7")', 'ws.Sort.Header = xlNo', 'ws.Sort.MatchCase = True', 'ws.Sort.Apply'],
     'A1:A7', 'case.xlsx'),
    (['ws.Range("A1:A7").Value = ' + VALUES,
      'ws.Range("A1:A7").Sort Key1:=ws.Range("A1"), Order1:=xlAscending, Header:=xlNo, OrderCustom:=n + 2, '
      'MatchCase:=True'], 'A1:A7', 'range_case.xlsx'),
    (['ws.Range("A1:A7").Value = ' + VALUES,
      'ws.Range("A1:A7").Sort Key1:=ws.Range("A1"), Order1:=xlAscending, Header:=xlNo, OrderCustom:=n + 2'],
     'A1:A7', 'range.xlsx'),
    (['ws.Range("A1:B3").Value = Application.Transpose(Application.Transpose(Array("x", "y", "z")))',
      'ws.Range("A1").Value = "t1"', 'ws.Range("A2").Value = 7', 'ws.Range("A3").Value = "t3"',
      'Application.AddCustomList ws.Range("A1:A3")',
      'ws.Range("C1").Value = Join(Application.GetCustomListContents(Application.CustomListCount), "|")',
      'ws.Range("C2").Value = Application.CustomListCount - n'], 'C1:C2', ''),
    (['ws.Range("A1:A7").Value = ' + VALUES, 'ws.Sort.SortFields.Clear',
      'ws.Sort.SortFields.Add Key:=ws.Range("A1:A7"), Order:=xlAscending, CustomOrder:=""',
      'ws.Sort.SetRange ws.Range("A1:A7")', 'ws.Sort.Header = xlNo', 'ws.Sort.Apply'], 'A1:A7', ''),
    (['ws.Range("A1:A7").Value = ' + VALUES, 'ws.Sort.SortFields.Clear',
      'ws.Sort.SortFields.Add Key:=ws.Range("A1:A7"), Order:=xlAscending, CustomOrder:=n + 50',
      'ws.Sort.SetRange ws.Range("A1:A7")', 'ws.Sort.Header = xlNo', 'ws.Sort.Apply'], 'A1:A7', ''),
]


def _probe() -> str:
    lines = ['Private Function Column(ByVal ws As Worksheet, ByVal address As String) As String',
             '    Dim c As Range, s As String', '    For Each c In ws.Range(address)',
             '        s = s & "|" & c.Formula', '    Next', '    Column = Mid(s, 2)', 'End Function', '',
             'Private Function Tried(ByVal which As Long, ByVal folder As String, ByVal n As Long) As String',
             '    Dim wb As Workbook, ws As Worksheet', '    Set wb = Workbooks.Add(xlWBATWorksheet)',
             '    Set ws = wb.Worksheets(1)', '    On Error GoTo Failed', '    Select Case which']
    for index, (statements, read, saved) in enumerate(CASES):
        lines.append(f'    Case {index}')
        lines += [f'        {statement}' for statement in statements]
        lines.append(f'        Tried = Column(ws, "{read}")')
        if saved:
            lines.append(f'        wb.SaveAs folder & "\\{saved}", xlOpenXMLWorkbook')
    lines += ['    End Select', 'Done:', '    wb.Close False', '    Exit Function', 'Failed:',
              '    Tried = "Err " & Err.Number & " " & Err.Description', '    Resume Done', 'End Function', '',
              'Public Function Probe(ByVal folder As String) As String', '    Dim s As String, i As Long, n As Long',
              '    Application.DisplayAlerts = False', '    n = Application.CustomListCount',
              '    Application.AddCustomList Array("Low", "Medium", "High")',
              f'    For i = 0 To {len(CASES) - 1}', '        s = s & "@@@" & Tried(i, folder, n)', '    Next',
              '    For i = Application.CustomListCount To n + 1 Step -1', '        Application.DeleteCustomList i',
              '    Next', '    Application.DisplayAlerts = True', '    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def main() -> None:
    source = _probe()
    with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(source, 'Probe', (folder,), timeout=300.0)
        assert result.ok and not result.dialogs, result
        answers = str(result.value).split('@@@')
        parts = {}
        for _, _, saved in CASES:
            if saved:
                with zipfile.ZipFile(Path(folder) / saved) as package:
                    sheet = package.read('xl/worksheets/sheet1.xml').decode('utf-8')
                found = re.search(r'<sortState\b.*?(?:</sortState>|/>)', sheet, re.DOTALL)
                parts[saved] = found.group(0) if found else ''
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source,
              'cases': [[statements, read, saved, answer]
                        for (statements, read, saved), answer in zip(CASES, answers, strict=True)],
              'parts': parts}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for statements, _, saved, answer in record['cases']:
        print(statements[-1][:70], '->', answer, '|', parts.get(saved, ''))


if __name__ == '__main__':
    main()
