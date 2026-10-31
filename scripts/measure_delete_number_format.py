"""Measure Workbook.DeleteNumberFormat: what it removes, what cells using the format become, and what it refuses.

Each case runs on a new workbook and answers as text; some save it, and the numFmts its styles.xml holds and the
format cell A1's xf points at are recorded.

    python scripts/measure_delete_number_format.py

writes tests/fixtures/delete_number_format.json, which tests/test_excel_delete_number_format.py replays.
"""
from __future__ import annotations

import json
import re
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/delete_number_format.json'
CUSTOM = '"0.000\\x"'
#: Each case: statements on a new workbook wb with sheet ws; the expression read back; the file saved, if any.
CASES = [
    (['ws.Range("A1").Value = 1.5', f'ws.Range("A1").NumberFormat = {CUSTOM}', f'wb.DeleteNumberFormat {CUSTOM}',
      'r = ws.Range("A1").NumberFormat & "|" & ws.Range("A1").Text'], 'r', 'used.xlsx'),
    ([f'ws.Range("A1").NumberFormat = {CUSTOM}', 'ws.Range("A1").NumberFormat = "General"',
      f'wb.DeleteNumberFormat {CUSTOM}', 'r = "deleted"'], 'r', 'unused.xlsx'),
    ([f'wb.DeleteNumberFormat {CUSTOM}', 'r = "deleted"'], 'r', ''),
    (['wb.DeleteNumberFormat "0.00"'], 'r', ''),
    (['wb.DeleteNumberFormat "General"'], 'r', ''),
    (['wb.DeleteNumberFormat "m/d/yyyy"'], 'r', ''),
    (['wb.DeleteNumberFormat ""'], 'r', ''),
    (['ws.Range("A1").Value = 1.5', 'ws.Range("A1").NumberFormat = "0.0000"', 'wb.DeleteNumberFormat "0.0000"',
      'r = ws.Range("A1").NumberFormat & "|" & ws.Range("A1").Text'], 'r', ''),
    (['ws.Range("A1").Value = 1.5', 'ws.Range("A1").NumberFormat = "0.0000"', 'wb.DeleteNumberFormat "0.0000 "',
      'r = ws.Range("A1").NumberFormat'], 'r', ''),
    (['ws.Range("A1").Value = 1.5', 'ws.Range("A1").NumberFormat = "0.0000"', 'wb.DeleteNumberFormat "0.0000;"',
      'r = ws.Range("A1").NumberFormat'], 'r', ''),
    (['ws.Range("A1").Value = 1.5', 'ws.Range("A1").NumberFormat = "[Red]0.0000"',
      'wb.DeleteNumberFormat "[red]0.0000"', 'r = ws.Range("A1").NumberFormat'], 'r', ''),
    (['ws.Range("A1").Value = 1.5', 'ws.Range("A1").NumberFormat = "0.0000"', 'wb.DeleteNumberFormat 5',
      'r = ws.Range("A1").NumberFormat'], 'r', ''),
    (['ws.Range("A1").Value = 1.5', 'ws.Range("A1").NumberFormat = "#,##0.0000"',
      'ws.Range("B1").NumberFormat = "#,##0.0000"', 'wb.Worksheets.Add.Range("C1").NumberFormat = "#,##0.0000"',
      'wb.DeleteNumberFormat "#,##0.0000"',
      'r = ws.Range("A1").NumberFormat & "|" & ws.Range("B1").NumberFormat & "|" & '
      'wb.Worksheets(1).Range("C1").NumberFormat'], 'r', 'several.xlsx'),
    (['ws.Range("A1").Value = 1.5', 'ws.Range("A1").NumberFormat = "0.0000"', 'ws.Range("A1").Font.Bold = True',
      'wb.DeleteNumberFormat "0.0000"', 'r = ws.Range("A1").NumberFormat & "|" & ws.Range("A1").Font.Bold'], 'r', ''),
    (['ws.Range("A1").Value = 1.5', 'ws.Range("A1").NumberFormat = "0.0000"', 'wb.Styles.Add "S1"',
      'wb.Styles("S1").NumberFormat = "0.0000"', 'wb.DeleteNumberFormat "0.0000"',
      'r = wb.Styles("S1").NumberFormat'], 'r', ''),
    (['ws.Range("A1").Value = 1.5', 'ws.Range("A1").NumberFormat = "0.0000"', 'ws.Protect',
      'wb.DeleteNumberFormat "0.0000"', 'r = ws.Range("A1").NumberFormat'], 'r', ''),
    (['ws.Range("A1").Value = 1.5', 'ws.Range("A1").NumberFormat = "0.0000"', 'wb.Protect Structure:=True',
      'wb.DeleteNumberFormat "0.0000"', 'r = ws.Range("A1").NumberFormat'], 'r', ''),
]


def _probe() -> str:
    lines = ['Private Function Tried(ByVal which As Long, ByVal folder As String) As String',
             '    Dim wb As Workbook, ws As Worksheet, r As Variant', '    Set wb = Workbooks.Add(xlWBATWorksheet)',
             '    Set ws = wb.Worksheets(1)', '    r = "ok"', '    On Error GoTo Failed', '    Select Case which']
    for index, (statements, read, saved) in enumerate(CASES):
        lines.append(f'    Case {index}')
        lines += [f'        {statement}' for statement in statements]
        lines.append(f'        Tried = CStr({read})')
        if saved:
            lines.append(f'        wb.SaveAs folder & "\\{saved}", xlOpenXMLWorkbook')
    lines += ['    End Select', 'Done:', '    wb.Close False', '    Exit Function', 'Failed:',
              '    Tried = "Err " & Err.Number & " " & Err.Description', '    Resume Done', 'End Function', '',
              'Public Function Probe(ByVal folder As String, ByVal first As Long, ByVal last As Long) As String',
              '    Dim s As String, i As Long', '    Application.DisplayAlerts = False', '    For i = first To last',
              '        s = s & "@@@" & Tried(i, folder)', '    Next', '    Application.DisplayAlerts = True',
              '    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def main() -> None:
    source = _probe()
    with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(source, 'Probe', (folder, 0, len(CASES) - 1), timeout=300.0)
        assert result.ok and not result.dialogs, result
        answers = str(result.value).split('@@@')
        parts = {}
        for _, _, saved in CASES:
            if saved and (Path(folder) / saved).exists():
                with zipfile.ZipFile(Path(folder) / saved) as package:
                    styles = package.read('xl/styles.xml').decode('utf-8')
                found = re.search(r'<numFmts\b.*?</numFmts>', styles, re.DOTALL)
                parts[saved] = found.group(0) if found else ''
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source,
              'cases': [[statements, read, saved, answer]
                        for (statements, read, saved), answer in zip(CASES, answers, strict=True)],
              'parts': parts}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for statements, _, saved, answer in record['cases']:
        print(statements[-1][:70], '->', answer, '|', parts.get(saved, '')[:120])


if __name__ == '__main__':
    main()
