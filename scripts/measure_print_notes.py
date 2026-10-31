"""PageSetup.PrintNotes: what it reads beside PrintComments, what a set does to PrintComments, and the pageSetup a
save writes after it.

    python scripts/measure_print_notes.py

writes tests/fixtures/print_notes.json, which tests/test_excel_print_notes.py replays.
"""
from __future__ import annotations

import json
import re
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/print_notes.json'
#: Each case: statements on a new workbook's sheet ws, and the file saved, if any.
CASES = [
    ([], ''),
    (['ws.PageSetup.PrintNotes = True'], 'on.xlsx'),
    (['ws.PageSetup.PrintNotes = True', 'ws.PageSetup.PrintNotes = False'], 'off.xlsx'),
    (['ws.PageSetup.PrintComments = xlPrintSheetEnd'], ''),
    (['ws.PageSetup.PrintComments = xlPrintInPlace'], ''),
    (['ws.PageSetup.PrintComments = xlPrintInPlace', 'ws.PageSetup.PrintNotes = False'], ''),
    (['ws.PageSetup.PrintComments = xlPrintInPlace', 'ws.PageSetup.PrintNotes = True'], ''),
    (['ws.PageSetup.PrintNotes = 7'], ''),
    (['ws.PageSetup.PrintNotes = "True"'], ''),
    (['ws.PageSetup.PrintNotes = "x"'], ''),
    (['ws.PageSetup.PrintComments = xlPrintSheetEnd', 'ws.PageSetup.PrintComments = xlPrintNoComments'],
     'no_comments.xlsx'),
    (['ws.PageSetup.PrintErrors = xlPrintErrorsDash', 'ws.PageSetup.PrintErrors = xlPrintErrorsDisplayed'],
     'errors.xlsx'),
    (['ws.PageSetup.Order = xlOverThenDown', 'ws.PageSetup.Order = xlDownThenOver'], 'order.xlsx'),
    (['ws.PageSetup.PrintComments = xlPrintNoComments'], 'never_set.xlsx'),
]


def _probe() -> str:
    lines = ['Private Function Tried(ByVal which As Long, ByVal folder As String) As String',
             '    Dim wb As Workbook, ws As Worksheet', '    Set wb = Workbooks.Add(xlWBATWorksheet)',
             '    Set ws = wb.Worksheets(1)', '    On Error GoTo Failed', '    Select Case which']
    for index, (statements, saved) in enumerate(CASES):
        lines.append(f'    Case {index}')
        lines += [f'        {statement}' for statement in statements]
        if saved:
            lines.append(f'        wb.SaveAs folder & "\\{saved}", xlOpenXMLWorkbook')
    lines += ['    End Select', '    Tried = ws.PageSetup.PrintNotes & " " & ws.PageSetup.PrintComments', 'Done:',
              '    wb.Close False', '    Exit Function', 'Failed:', '    Tried = "Err " & Err.Number', '    Resume Done',
              'End Function', '', 'Public Function Probe(ByVal folder As String) As String',
              '    Dim s As String, i As Long', '    Application.DisplayAlerts = False',
              f'    For i = 0 To {len(CASES) - 1}', '        s = s & "@@@" & Tried(i, folder)', '    Next',
              '    Application.DisplayAlerts = True', '    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def main() -> None:
    source = _probe()
    with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(source, 'Probe', (folder,), timeout=300.0)
        assert result.ok and not result.dialogs, result
        answers = str(result.value).split('@@@')
        parts = {}
        for _, saved in CASES:
            if saved:
                with zipfile.ZipFile(Path(folder) / saved) as package:
                    sheet = package.read('xl/worksheets/sheet1.xml').decode('utf-8')
                found = re.search(r'<pageSetup\b[^>]*>', sheet)
                parts[saved] = found.group(0) if found else ''
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source,
              'cases': [[statements, saved, answer] for (statements, saved), answer in zip(CASES, answers, strict=True)],
              'parts': parts}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for statements, saved, answer in record['cases']:
        print(statements, '->', answer, '|', parts.get(saved, ''))


if __name__ == '__main__':
    main()
