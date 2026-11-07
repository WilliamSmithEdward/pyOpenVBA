"""Ranges held in variables, and the same references in a formula, across a row inserted at 1 and then a column inserted at A, when
they reach the sheet's last row or column. Rival rules: the top or left moves and the far end stays at the sheet's edge, so ws.Cells
becomes 2:1048576, as the model had it, which held-out round 51106 refuted once; or a reference spanning every row (or every column)
stays whole while one that only reaches the edge moves its near end. Each is read on a new book and again after Excel's own save and
reopen.

    python scripts/measure_whole_span_inserts.py

Takes excel-coordination/live.lock. Writes tests/fixtures/whole_span_inserts.json, which tests/test_excel_whole_span_inserts.py replays.
"""
from __future__ import annotations

import json
import os
import subprocess
import tempfile
import time
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / 'tests' / 'fixtures' / 'whole_span_inserts.json'
#: Held as ranges, and named in Z1's formula in the same order.
REFS = ['A1:XFD1048576', 'A:B', '1:3', 'A1:B1048576', 'B2:C1048576', 'A1:XFD3', 'B2:XFD4', 'C5']


def source() -> str:
    held = ', '.join(f'r{at} As Object' for at in range(len(REFS)))
    sets = ': '.join(f'Set r{at} = ws.Range("{ref}")' for at, ref in enumerate(REFS))
    shown = ' & " " & '.join(f'r{at}.Address(False, False)' for at in range(len(REFS)))
    formula = '+'.join(f'ROWS({ref})' for ref in REFS)
    return (f'Public Function Probe(ByVal path As String, ByVal reopen As Boolean) As String\n    Dim wb As Object, ws As Object, {held}, s As String\n'
            '    Application.DisplayAlerts = False\n    Set wb = Workbooks.Add(xlWBATWorksheet)\n    Set ws = wb.Worksheets(1)\n'
            f'    ws.Range("Z1").Formula = "={formula}"\n'
            '    If reopen Then\n        wb.SaveAs path, xlOpenXMLWorkbook\n        wb.Close False\n        Set wb = Workbooks.Open(path)\n'
            f'        Set ws = wb.Worksheets(1)\n    End If\n    {sets}\n'
            f'    ws.Rows(1).Insert\n    s = "rows " & {shown} & " | " & ws.Range("Z2").Formula\n'
            f'    ws.Columns(1).Insert\n    s = s & " || columns " & {shown} & " | " & ws.Range("AA2").Formula\n'
            '    wb.Close False\n    Probe = s\nEnd Function\n')


def _lock() -> Path:
    common = subprocess.run(['git', 'rev-parse', '--git-common-dir'], capture_output=True, text=True, check=True, cwd=ROOT).stdout.strip()
    return (ROOT / common).resolve() / 'excel-coordination' / 'live.lock'


def main() -> None:
    from pyvbaharness import ExcelSession, HarnessConfig

    found: list[dict[str, object]] = []
    lock = _lock()
    started = time.monotonic()
    while True:
        try:
            os.mkdir(lock)
            break
        except FileExistsError:
            if time.monotonic() - started > 420:
                raise SystemExit('live.lock is still held by another session; try again later') from None
            time.sleep(10)
    try:
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as held, ExcelSession(HarnessConfig(lock_wait_s=1200.0, exclusive=False)) as excel:
            excel.new_document()
            for reopen in (False, True):
                done = excel.run_vba(source(), 'Probe', (str(Path(held) / 'w.xlsx'), reopen), timeout=60.0, line_numbers=False)
                assert done.ok, done
                found.append({'reopen': reopen, 'answer': str(done.value)})
                print(reopen, done.value, flush=True)
    finally:
        os.rmdir(lock)
    OUT.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'refs': REFS, 'source': source(), 'cases': found}, indent=1) + '\n', encoding='utf-8')


if __name__ == '__main__':
    main()
