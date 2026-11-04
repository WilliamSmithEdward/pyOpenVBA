"""Range.ColumnWidth and RowHeight given each kind of value: only the error, late and early bound, and whether the size
changed. The size itself follows the display's pixels and is measured at 96 DPI by scripts/measure_dimensions.py, so
nothing here records it, and this runs on any display.

    python scripts/measure_range_sizes.py

writes tests/fixtures/range_sizes.json, which tests/test_excel_range_sizes.py replays.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/range_sizes.json'
#: The bounds are in the unit each takes, far enough from a pixel's rounding that the display cannot decide them.
VALUES = {'ColumnWidth': ['20', '0', '100', '255', '256', '-1', '1.5', '"20"', '"x"', '""', 'Null', 'Empty', 'True', 'False', 'Application'],
          'RowHeight': ['20', '0', '100', '409', '410', '-1', '1.5', '"20"', '"x"', '""', 'Null', 'Empty', 'True', 'False', 'Application']}
TARGETS = ['ws.Range("B2")', 'ws.Range("B2:C3")', 'ws.Columns(2)', 'ws.Rows(2)']


def cases() -> list[tuple[str, str, str]]:
    return [(member, target, value) for member, values in VALUES.items() for target in TARGETS for value in values]


def source(typed: bool = False) -> str:
    kind = 'Range' if typed else 'Object'
    body = 'Private wb As Object\n\nPublic Sub Setup()\n    Set wb = Workbooks.Add(xlWBATWorksheet)\nEnd Sub\n'
    for number, (member, target, value) in enumerate(cases()):
        body += (f'\nPublic Function S{number}() As String\n    Dim ws As Object, c As {kind}, before As Variant, after As Variant, e As Long\n'
                 f'    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))\n    Set c = {target}\n'
                 f'    On Error Resume Next\n    before = c.{member}\n    Err.Clear\n    c.{member} = {value}\n    e = Err.Number\n    Err.Clear\n'
                 f'    after = c.{member}\n    S{number} = "e" & e & "|" & IIf(before = after, "same", "changed")\nEnd Function\n')
    return body


def _probe(typed: bool) -> str:
    calls = ''.join(f'    out = out & "@@" & S{number}()\n' for number in range(len(cases())))
    return (source(typed) + f'\nPublic Function Probe() As String\n    Dim out As String\n    Setup\n{calls}'
            '    Probe = Mid(out, 3)\n    wb.Close False\nEnd Function\n')


def main() -> None:
    every = cases()
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0, exclusive=False)) as excel:
        excel.new_document()
        late = excel.run_vba(_probe(False), 'Probe', (), timeout=200.0)
        assert late.ok and not late.dialogs, late
        early = excel.run_vba(_probe(True), 'Probe', (), timeout=200.0)
        assert early.ok and not early.dialogs, early
        standard = excel.run_vba('Public Function Probe() As Double\n    Probe = ActiveSheet.StandardHeight\nEnd Function\n', 'Probe', (), timeout=30.0)
    answers, bound = str(late.value).split('@@'), str(early.value).split('@@')
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'standard_height': standard.value, 'source': source(),
              'early': {str(number): bound[number] for number in range(len(every)) if bound[number] != answers[number]},
              'cases': [[*every[number], answers[number]] for number in range(len(every))]}
    OUT.write_text(json.dumps(record, indent=0) + '\n', encoding='utf-8')
    print(len(every), 'cases', len(record['early']), 'differ typed; StandardHeight', standard.value)


if __name__ == '__main__':
    main()
