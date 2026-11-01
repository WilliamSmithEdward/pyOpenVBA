"""Window.ActivateNext and Window.ActivatePrevious: which window each makes active, from where, and what each answers.

Each case opens three new workbooks, labelled a, b and c in the order they were added, so c's window is active, runs
the case's statements, and answers the active workbook's label after each step: a, b or c, or "other" for a window of
a workbook the case did not make, such as the one the harness runs in. Windows.Count is recorded as well, as the
windows the steps go round.

    python scripts/measure_window_activate_next.py

writes tests/fixtures/window_activate_next.json, which tests/test_excel_window_activate_next.py replays.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/window_activate_next.json'
STEP = 's = s & "," & Label()'
CASES = [
    f'ActiveWindow.ActivateNext: {STEP}',
    f'ActiveWindow.ActivatePrevious: {STEP}',
    f'ActiveWindow.ActivateNext: {STEP}: ActiveWindow.ActivateNext: {STEP}: ActiveWindow.ActivateNext: {STEP}',
    f'ActiveWindow.ActivatePrevious: {STEP}: ActiveWindow.ActivatePrevious: {STEP}: ActiveWindow.ActivatePrevious: '
    f'{STEP}',
    f'ActiveWindow.ActivateNext: {STEP}: ActiveWindow.ActivatePrevious: {STEP}',
    f'books(1).Activate: {STEP}: ActiveWindow.ActivateNext: {STEP}: ActiveWindow.ActivatePrevious: {STEP}',
    f'books(1).Windows(1).ActivateNext: {STEP}',
    f's = s & "," & TypeName(ActiveWindow.ActivateNext): {STEP}',
]


def _probe() -> str:
    lines = ['Private books(1 To 3) As Workbook', '',
             'Private Function Label() As String',
             '    Dim i As Long', '    Label = "other"',
             '    For i = 1 To 3',
             '        If ActiveWorkbook Is books(i) Then Label = Chr(96 + i)',
             '    Next', 'End Function', '',
             'Private Function Tried(ByVal which As Long) As String',
             '    Dim i As Long, s As String',
             '    For i = 1 To 3: Set books(i) = Workbooks.Add(xlWBATWorksheet): Next',
             '    s = Label()',
             '    On Error GoTo Failed', '    Select Case which']
    for index, case in enumerate(CASES):
        lines += [f'    Case {index}', f'        {case}']
    lines += ['    End Select', '    Tried = "ok " & s', 'Done:',
              '    For i = 1 To 3: books(i).Close False: Next', '    Exit Function', 'Failed:',
              '    Tried = "Err " & Err.Number & " " & s', '    Resume Done', 'End Function', '',
              'Public Function Probe() As String', '    Dim i As Long, s As String',
              '    s = "windows=" & (Windows.Count + 3)',
              f'    For i = 0 To {len(CASES) - 1}', '        s = s & "@@@" & Tried(i)', '    Next',
              '    Probe = s', 'End Function']
    return '\n'.join(lines) + '\n'


def main() -> None:
    source = _probe()
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(source, 'Probe', timeout=280.0)
        assert result.ok and not result.dialogs, result
        answers = str(result.value).split('@@@')
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source, 'windows': answers[0],
              'cases': [list(pair) for pair in zip(CASES, answers[1:], strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    print(answers[0])
    for case, answer in record['cases']:
        print(case[:70], '->', answer)


if __name__ == '__main__':
    main()
