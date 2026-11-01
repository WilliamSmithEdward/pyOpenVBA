"""Worksheet.EnableFormatConditionsCalculation: what it reads after each setting, the sheetPr the saved sheet holds,
and what it reads once the workbook is opened again.

Each case makes a workbook, sets the property on its sheet as the case says, reads it, saves the workbook, opens it
again and reads it once more.

    python scripts/measure_format_conditions_calculation.py

writes tests/fixtures/format_conditions_calculation.json, which tests/test_excel_format_conditions_calculation.py
replays.
"""
from __future__ import annotations

import json
import re
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/format_conditions_calculation.json'
CASES = [
    '',
    'ws.EnableFormatConditionsCalculation = False',
    'ws.EnableFormatConditionsCalculation = False: ws.EnableFormatConditionsCalculation = True',
    'ws.EnableFormatConditionsCalculation = 0',
    'ws.EnableFormatConditionsCalculation = "False"',
    'ws.EnableFormatConditionsCalculation = Null',
    'ws.EnableFormatConditionsCalculation = "abc"',
    'Application.Calculation = xlCalculationManual: ws.EnableFormatConditionsCalculation = False',
]


def _probe() -> str:
    lines = ['Private Function Tried(ByVal which As Long, ByVal folder As String) As String',
             '    Dim wb As Workbook, ws As Object, s As String, path As String, calc As Long',
             '    calc = Application.Calculation', '    Set wb = Workbooks.Add(xlWBATWorksheet)',
             '    Set ws = wb.Worksheets(1)', '    ws.Range("A1").Value = 1', '    path = folder & "\\c" & which & ".xlsx"',
             '    On Error GoTo Failed', '    Select Case which']
    for index, case in enumerate(CASES):
        lines += [f'    Case {index}'] + ([f'        {case}'] if case else [])
    lines += ['    End Select', '    s = "ok|" & ws.EnableFormatConditionsCalculation', 'Saved:',
              '    Application.Calculation = calc', '    wb.SaveAs path, xlOpenXMLWorkbook', '    wb.Close False',
              '    Set wb = Workbooks.Open(path)', '    s = s & "|" & wb.Worksheets(1).EnableFormatConditionsCalculation',
              '    wb.Close False', '    Tried = s', '    Exit Function', 'Failed:',
              '    s = "Err " & Err.Number', '    Resume Saved', 'End Function', '',
              'Public Function Probe(ByVal folder As String) As String', '    Dim i As Long, s As String',
              '    Application.DisplayAlerts = False', f'    For i = 0 To {len(CASES) - 1}',
              '        s = s & "@@@" & Tried(i, folder)', '    Next', '    Application.DisplayAlerts = True',
              '    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def sheet_properties(path: Path) -> str:
    with zipfile.ZipFile(path) as package:
        xml = package.read('xl/worksheets/sheet1.xml').decode('utf-8')
    found = re.search(r'<sheetPr\b.*?(?:/>|</sheetPr>)', xml, re.S)
    return found.group() if found else ''


def main() -> None:
    source = _probe()
    with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(source, 'Probe', (folder,), timeout=280.0)
        assert result.ok and not result.dialogs, result
        answers = str(result.value).split('@@@')
        parts = [sheet_properties(Path(folder) / f'c{index}.xlsx') for index in range(len(CASES))]
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source,
              'cases': [[case, answer, part] for case, answer, part in zip(CASES, answers, parts, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for case, answer, part in record['cases']:
        print(case, '->', answer, '|', part)


if __name__ == '__main__':
    main()
