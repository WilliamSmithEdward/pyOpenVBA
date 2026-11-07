"""Collections a macro takes from a range and holds while rows go in above it: Range.Hyperlinks, Range.FormatConditions, and a
FormatCondition reached through D3. Rival rules: the collection keeps the place it was taken from, as the model had it, which
held-out round 51103 refuted once for Hyperlinks; or it follows its cells, as a held Range does. Each is read on a new book and
again after Excel's own save and reopen. Answers carry TypeName.

    python scripts/measure_held_range_collections.py

Takes excel-coordination/live.lock. Writes tests/fixtures/held_range_collections.json, which
tests/test_excel_held_range_collections.py replays.
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
OUT = ROOT / 'tests' / 'fixtures' / 'held_range_collections.json'
#: Held first, then two rows go in above, then read.
CASES = {
    'links held': ('Set h = ws.Range("Q2").Hyperlinks', 'ws.Rows("1:2").Insert', 'x = h.Count & "/" & h.Item(1).Range.Address(False, False)'),
    'links of a held range': ('Set r = ws.Range("Q2")', 'ws.Rows("1:2").Insert', 'x = r.Hyperlinks.Count & "/" & r.Hyperlinks(1).Range.Address(False, False)'),
    'links held, link added after': ('Set h = ws.Range("Q2:Q3").Hyperlinks', 'ws.Rows("1:2").Insert: ws.Hyperlinks.Add ws.Range("Q5"), "http://example.com/b"',
                                     'x = h.Count & "/" & h.Item(h.Count).Range.Address(False, False)'),
    'links held, moved column': ('Set h = ws.Range("Q2").Hyperlinks', 'ws.Columns(1).Insert', 'x = h.Count & "/" & h.Item(1).Range.Address(False, False)'),
    'formats held': ('Set h = ws.Range("D1:D5").FormatConditions', 'ws.Rows("1:2").Insert',
                     'x = h.Count & "/" & h.Item(1).Parent.Address(False, False) & "/" & h.Item(1).AppliesTo.Address(False, False)'),
    'format held through D3': ('Set h = ws.Range("D3").FormatConditions(1)', 'ws.Rows("1:2").Insert',
                               'x = h.Parent.Address(False, False) & "/" & h.AppliesTo.Address(False, False)'),
    'format held, then rows deleted under it': ('Set h = ws.Range("D3").FormatConditions(1)', 'ws.Rows("3").Delete',
                                                'x = h.Parent.Address(False, False) & "/" & h.AppliesTo.Address(False, False)'),
}


def source() -> str:
    text = ('Public Function Probe(ByVal path As String, ByVal which As Long, ByVal reopen As Boolean) As String\n'
            '    Dim wb As Object, ws As Object, h As Object, r As Object, x As String\n    Application.DisplayAlerts = False\n'
            '    Set wb = Workbooks.Add(xlWBATWorksheet)\n    Set ws = wb.Worksheets(1)\n    ws.Range("D1:D5").Value = 3\n'
            '    ws.Range("D1:D5").FormatConditions.Add xlCellValue, xlGreater, "=2"\n'
            '    ws.Hyperlinks.Add ws.Range("Q2"), "http://example.com/a", "", "", "L0"\n'
            '    If reopen Then\n        wb.SaveAs path, xlOpenXMLWorkbook\n        wb.Close False\n        Set wb = Workbooks.Open(path)\n'
            '        Set ws = wb.Worksheets(1)\n    End If\n    On Error Resume Next\n    Select Case which\n')
    for at, (hold, edit, read) in enumerate(CASES.values()):
        text += (f'    Case {at}\n        {hold}\n        If Err.Number <> 0 Then x = "hold E" & Err.Number: Err.Clear\n        {edit}\n'
                 f'        If Err.Number <> 0 Then x = x & "edit E" & Err.Number: Err.Clear\n        {read}\n')
    return text + '    End Select\n    Probe = x & "|E" & Err.Number\n    Err.Clear\n    wb.Close False\nEnd Function\n'


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
            for at, name in enumerate(CASES):
                for reopen in (False, True):
                    done = excel.run_vba(source(), 'Probe', (str(Path(held) / f'h{at}.xlsx'), at, reopen), timeout=60.0, line_numbers=False)
                    assert done.ok, done
                    found.append({'case': name, 'reopen': reopen, 'answer': str(done.value)})
                    print(name, reopen, done.value, flush=True)
    finally:
        os.rmdir(lock)
    OUT.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source(), 'cases': found}, indent=1) + '\n', encoding='utf-8')


if __name__ == '__main__':
    main()
