"""Where an error raised inside a macro that Application.Run started goes: to the caller's On Error, or to Excel's run-time error
dialog past every handler.

Each case is its own run. Probe calls a procedure of the module directly or through Application.Run, the callee raises or does not,
and Probe returns what its handler saw. The harness answers a run-time error dialog with End, which ends every macro, so a case that
shows one records the dialog's error number and nothing else. The model's answers are written first.

    python scripts/measure_run_errors.py predict
    python scripts/measure_run_errors.py measure

writes tests/fixtures/vba_run_errors.json, which tests/test_vba_run_errors.py replays. Answers are the value Probe returned,
"error N" for an error that reached the host past Probe, and "dialog N" for Excel's run-time error dialog.
"""
from __future__ import annotations

import hashlib
import json
import re
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / 'tests/fixtures/vba_run_errors.json'
CALLEES = '''Public Sub Inner()
    Err.Raise 5
End Sub
Public Sub Fine()
End Sub
Public Sub Divide()
    Dim x As Double
    x = 1 / 0
End Sub
Public Function Twice(ByVal n As Long) As Long
    If n < 0 Then Err.Raise 6
    Twice = 2 * n
End Function
Public Sub Handled()
    On Error Resume Next
    Err.Raise 5
End Sub
Public Sub Deeper()
    Err.Raise 7
End Sub
Public Sub Outer()
    On Error GoTo h
    Deeper
    Exit Sub
h:
End Sub
Public Sub Unhandled()
    Deeper
End Sub
Public Sub Custom()
    Err.Raise vbObjectError + 513, "Probe", "custom"
End Sub
Public Sub InHandler()
    On Error GoTo h
    Err.Raise 5
    Exit Sub
h:
    Err.Raise 9
End Sub
'''
RESUME = 'Function Probe() As String\n    On Error Resume Next\n    @@\n    Probe = "after " & Err.Number\nEnd Function\n'
GOTO = ('Function Probe() As String\n    On Error GoTo h\n    @@\n    Probe = "no error"\n    Exit Function\nh:\n'
        '    Probe = "trapped " & Err.Number\nEnd Function\n')
BARE = 'Function Probe() As String\n    @@\n    Probe = "no error"\nEnd Function\n'
CASES: list[tuple[str, str, str]] = [
    ('direct resume', RESUME, 'Inner'), ('direct goto', GOTO, 'Inner'), ('run resume', RESUME, 'Application.Run "Inner"'),
    ('run goto', GOTO, 'Application.Run "Inner"'), ('run bare', BARE, 'Application.Run "Inner"'), ('run fine', GOTO, 'Application.Run "Fine"'),
    ('run divide', RESUME, 'Application.Run "Divide"'), ('run function', GOTO, 'Probe = Application.Run("Twice", 2): Exit Function'),
    ('run function raises', RESUME, 'Dim v As Variant: v = Application.Run("Twice", -1)'), ('run handled', RESUME, 'Application.Run "Handled"'),
    ('run handled deeper', GOTO, 'Application.Run "Outer"'), ('run unhandled deeper', RESUME, 'Application.Run "Unhandled"'),
    ('run custom', GOTO, 'Application.Run "Custom"'), ('run in handler', RESUME, 'Application.Run "InHandler"'),
    ('run missing', GOTO, 'Application.Run "Nowhere"'), ('run extra argument', RESUME, 'Application.Run "Fine", 1'),
    ('direct unhandled deeper', GOTO, 'Unhandled'), ('call run', RESUME, 'Call Application.Run("Inner")'),
]


def cases() -> list[dict[str, str]]:
    return [{'name': name, 'source': CALLEES + body.replace('@@', line)} for name, body, line in CASES]


def model(case: dict[str, str]) -> str:
    sys.path.insert(0, str(ROOT / 'src'))
    from pyopenvba.apps.excel import ExcelApplication
    from pyopenvba.exceptions import VBARuntimeError

    app = ExcelApplication()
    app.add_workbook()
    app.add_module(case['source'], name='Harness')
    try:
        return str(app.run('Harness.Probe'))
    except VBARuntimeError as caught:
        return f'error {caught.number}'


def predict() -> None:
    every = cases()
    text = json.dumps({'cases': every, 'model': [model(case) for case in every]}, indent=1) + '\n'
    OUT.write_text(text, encoding='utf-8', newline='\n')
    print(len(every), 'cases; predictions sha256', hashlib.sha256(text.encode('utf-8')).hexdigest())


def measure() -> None:
    from pyvbaharness import ExcelSession

    record: dict[str, Any] = json.loads(OUT.read_text(encoding='utf-8'))
    answers: list[str] = []
    with ExcelSession() as excel:
        for case in record['cases']:
            excel.new_document()
            result = excel.run_vba(case['source'], 'Probe', timeout=30, line_numbers=False)
            shown = [dialog for dialog in result.dialogs if dialog.classification == 'runtime-error']
            if shown:
                found = re.search(r"error '(-?\d+)'", shown[0].message)
                answers.append(f'dialog {found.group(1) if found else shown[0].message[:60]}')
            elif result.ok:
                answers.append(str(result.value))
            elif result.error is not None:
                answers.append(f'error {result.error.number}')
            else:
                answers.append(f'stopped {result.outcome} {result.message[:60]}')
            assert all(dialog.classification == 'runtime-error' for dialog in result.dialogs), (case['name'], result.dialogs)
            print(case['name'], answers[-1], flush=True)
    record.update({'measured_at': datetime.now(timezone.utc).isoformat(), 'excel': answers})
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8', newline='\n')
    print('model wrong', sum(mine != theirs for mine, theirs in zip(record['model'], answers)), 'of', len(answers))


if __name__ == '__main__':
    {'predict': predict, 'measure': measure}[sys.argv[1]]()
