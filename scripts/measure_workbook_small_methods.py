"""Workbook.EnableConnections, ConvertComments, CanCheckIn and ExclusiveAccess: what each answers
and the errors each raises, on a new workbook and on one saved to a file.

Each case makes a workbook (saving it first where the case says), runs the case's statements and records what they
answer. AddToFavorites, which writes to the user's Favorites folder, and NewWindow, which opens a window, are left
alone, as is ToggleFormsDesign, whose design mode stops the macro running.

    python scripts/measure_workbook_small_methods.py

writes tests/fixtures/workbook_small_methods.json, which tests/test_excel_workbook_small_methods.py replays.
"""
from __future__ import annotations

import json
import tempfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/workbook_small_methods.json'
#: (whether the workbook is saved first, statements)
CASES = [
    (False, 'wb.EnableConnections: v = "done"'),
    (False, 'v = TypeName(wb.EnableConnections)'),
    (False, 'wb.ConvertComments: v = "done"'),
    (False, 'v = CStr(wb.CanCheckIn)'),
    (True, 'v = CStr(wb.CanCheckIn)'),
    (False, 'v = CStr(wb.ExclusiveAccess)'),
    (True, 'v = CStr(wb.ExclusiveAccess)'),
    (True, 'wb.EnableConnections: v = "done"'),
]


def _probe() -> str:
    lines = ['Private Function Tried(ByVal which As Long, ByVal folder As String) As String',
             '    Dim wb As Object, ws As Object, v As Variant', '    Set wb = Workbooks.Add(xlWBATWorksheet)',
             '    Set ws = wb.Worksheets(1)', '    On Error GoTo Failed', '    Select Case which']
    for index, (save, statements) in enumerate(CASES):
        lines.append(f'    Case {index}')
        if save:
            lines.append(f'        wb.SaveAs folder & "\\c{index}.xlsx", xlOpenXMLWorkbook')
        lines.append(f'        {statements}')
    lines += ['    End Select', '    Tried = "ok " & v', 'Done:', '    wb.Close False', '    Exit Function', 'Failed:',
              '    Tried = "Err " & Err.Number', '    Resume Done', 'End Function', '',
              'Public Function Probe(ByVal folder As String) As String', '    Dim i As Long, s As String',
              '    Application.DisplayAlerts = False', f'    For i = 0 To {len(CASES) - 1}',
              '        s = s & "@@@" & Tried(i, folder)', '    Next', '    Application.DisplayAlerts = True',
              '    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def main() -> None:
    source = _probe()
    with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(source, 'Probe', (folder,), timeout=200.0)
        assert result.ok and not result.dialogs, result
        answers = str(result.value).split('@@@')
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source,
              'cases': [[f'{"saved" if save else "new"}: {statements}', answer]
                        for (save, statements), answer in zip(CASES, answers, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for case, answer in record['cases']:
        print(case, '->', answer)


if __name__ == '__main__':
    main()
