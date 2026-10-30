"""ReDim with an upper bound below its lower bound (issue #255): which shapes VBA refuses, and with what error.

Predictions written before measuring: every dimension whose upper bound is below its lower bound is error 9,
Preserve or not, and a dimension of one element (1 To 1) is fine.

    python scripts/measure_redim_bounds.py

writes tests/fixtures/redim_bounds.json, which tests/test_vba_redim_bounds.py replays.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/redim_bounds.json'
#: Each case: the statements on a dynamic array a() As Long, and on v As Variant.
CASES = [
    'ReDim a(1 To 0)',
    'ReDim a(0 To -1)',
    'ReDim a(5 To 3)',
    'ReDim a(1 To 1)',
    'ReDim a(-1)',
    'ReDim a(1 To 2, 1 To 0)',
    'ReDim a(1 To 0, 1 To 2)',
    'ReDim a(1 To 2): ReDim Preserve a(1 To 0)',
    'ReDim v(1 To 0)',
    'ReDim a(0)',
]


def _probe() -> str:
    lines = ['Private Function Tried(ByVal which As Long) As String', '    Dim a() As Long, v As Variant',
             '    On Error GoTo Failed', '    Select Case which']
    for index, statements in enumerate(CASES):
        lines += [f'    Case {index}', f'        {statements}']
    lines += ['    End Select', '    Tried = "ok"', '    Exit Function', 'Failed:',
              '    Tried = "Err " & Err.Number & " " & Err.Description', 'End Function', '',
              'Public Function Probe() As String', '    Dim s As String, i As Long',
              f'    For i = 0 To {len(CASES) - 1}', '        s = s & "|" & Tried(i)', '    Next',
              '    Probe = Mid(s, 2)', 'End Function']
    return '\n'.join(lines) + '\n'


def main() -> None:
    source = _probe()
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(source, 'Probe', timeout=120.0)
        assert result.ok and not result.dialogs, result
        answers = str(result.value).split('|')
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source,
              'cases': dict(zip(CASES, answers, strict=True))}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for case, answer in record['cases'].items():
        print(case, '->', answer)


if __name__ == '__main__':
    main()
