"""Workbook.RunAutoMacros: which of Auto_Open, Auto_Close, Auto_Activate and Auto_Deactivate each call runs, in one
sequence, and the errors.

The probe's own module defines Auto_Open, Auto_Close and Auto_Activate, each adding its name to a log, and no
Auto_Deactivate. One procedure runs the steps in order, clearing the log before each and answering what it holds
after, or the error number. Excel runs the macro on every other call that succeeds and does nothing on the calls
between, so the order is part of what is measured.

An error an auto macro raises is left out: the caller's On Error does not catch it, and Excel stops on its own
run-time error dialog instead.

    python scripts/measure_run_auto_macros.py

writes tests/fixtures/run_auto_macros.json, which tests/test_excel_run_auto_macros.py replays.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/run_auto_macros.json'
STEPS = [
    'ThisWorkbook.RunAutoMacros xlAutoOpen', 'ThisWorkbook.RunAutoMacros xlAutoClose',
    'ThisWorkbook.RunAutoMacros xlAutoClose', 'ThisWorkbook.RunAutoMacros xlAutoActivate',
    'ThisWorkbook.RunAutoMacros xlAutoActivate', 'ThisWorkbook.RunAutoMacros xlAutoDeactivate',
    'ThisWorkbook.RunAutoMacros 0', 'ThisWorkbook.RunAutoMacros xlAutoDeactivate',
    'ThisWorkbook.RunAutoMacros xlAutoOpen', 'ThisWorkbook.RunAutoMacros "1"',
    'ThisWorkbook.RunAutoMacros Which:=xlAutoClose', 'ThisWorkbook.RunAutoMacros Which:=xlAutoClose',
    'other.RunAutoMacros xlAutoOpen', 'ThisWorkbook.RunAutoMacros xlAutoOpen', 'ThisWorkbook.RunAutoMacros 5',
    'ThisWorkbook.RunAutoMacros -1', 'Application.Run "Auto_Open"', 'ThisWorkbook.RunAutoMacros xlAutoOpen',
    'ThisWorkbook.RunAutoMacros xlAutoOpen',
]


def _probe() -> str:
    lines = ['Private log As String', '',
             'Public Sub Auto_Open()', '    log = log & "open;"', 'End Sub', '',
             'Public Sub Auto_Close()', '    log = log & "close;"', 'End Sub', '',
             'Public Sub Auto_Activate()', '    log = log & "activate;"', 'End Sub', '',
             'Public Function Probe() As String', '    Dim s As String, other As Workbook',
             '    Set other = Workbooks.Add', '    On Error Resume Next']
    for step in STEPS:
        lines += ['    log = "": Err.Clear', f'    {step}',
                  '    s = s & "@@@" & IIf(Err.Number, "Err " & Err.Number, "ok " & log)']
    lines += ['    other.Close False', '    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def main() -> None:
    source = _probe()
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(source, 'Probe', timeout=200.0)
        assert result.ok and not result.dialogs, result
        answers = str(result.value).split('@@@')
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source,
              'cases': [list(pair) for pair in zip(STEPS, answers, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for case, answer in record['cases']:
        print(case, '->', answer)


if __name__ == '__main__':
    main()
