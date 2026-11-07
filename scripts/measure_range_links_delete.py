"""Range.Hyperlinks.Delete when the range has no link left. Rival rules: a range's links with none is error 1004, as the model had
it from held_objects.json, which held-out round 51104 refuted once on a range that never had one; or it is 1004 only for a held
collection whose links have all gone, and nothing for one that held none. Each is read on a new book and again after Excel's own
save and reopen, with the sheet's link count after.

    python scripts/measure_range_links_delete.py

Takes excel-coordination/live.lock. Writes tests/fixtures/range_links_delete.json, which tests/test_excel_range_links_delete.py replays.
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
OUT = ROOT / 'tests' / 'fixtures' / 'range_links_delete.json'
#: What happens before h.Delete, with links at D5 and D6 and none at Q2.
CASES = {
    'never had one': 'Set h = ws.Range("Q2").Hyperlinks',
    'never had one, whole column': 'Set h = ws.Range("Q:Q").Hyperlinks',
    'held, all taken': 'Set h = ws.Range("D5:D6").Hyperlinks: ws.Hyperlinks.Delete',
    'held, one of two taken': 'Set h = ws.Range("D5:D6").Hyperlinks: ws.Range("D5").Hyperlinks.Delete',
    'taken, then asked': 'ws.Range("D5:D6").Hyperlinks.Delete: Set h = ws.Range("D5:D6").Hyperlinks',
    'held, moved away': 'Set h = ws.Range("D5").Hyperlinks: ws.Columns(1).Insert',
    'asked after moving away': 'ws.Columns(1).Insert: Set h = ws.Range("D5").Hyperlinks',
}


def source() -> str:
    text = ('Public Function Probe(ByVal path As String, ByVal which As Long, ByVal reopen As Boolean) As String\n'
            '    Dim wb As Object, ws As Object, h As Object, s As String\n    Application.DisplayAlerts = False\n'
            '    Set wb = Workbooks.Add(xlWBATWorksheet)\n    Set ws = wb.Worksheets(1)\n'
            '    ws.Hyperlinks.Add ws.Range("D5"), "http://example.com/a"\n    ws.Hyperlinks.Add ws.Range("D6"), "http://example.com/b"\n'
            '    If reopen Then\n        wb.SaveAs path, xlOpenXMLWorkbook\n        wb.Close False\n        Set wb = Workbooks.Open(path)\n'
            '        Set ws = wb.Worksheets(1)\n    End If\n    On Error Resume Next\n    Select Case which\n')
    for at, before in enumerate(CASES.values()):
        text += f'    Case {at}\n        {before}\n'
    return text + ('    End Select\n    If Err.Number <> 0 Then s = "before E" & Err.Number & "|": Err.Clear\n    s = s & "count " & h.Count\n'
                   '    h.Delete\n    s = s & "|delete E" & Err.Number\n    Err.Clear\n    s = s & "|left " & ws.Hyperlinks.Count\n'
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
            for at, name in enumerate(CASES):
                for reopen in (False, True):
                    done = excel.run_vba(source(), 'Probe', (str(Path(held) / f'd{at}.xlsx'), at, reopen), timeout=60.0, line_numbers=False)
                    assert done.ok, done
                    found.append({'case': name, 'reopen': reopen, 'answer': str(done.value)})
                    print(name, reopen, done.value, flush=True)
    finally:
        os.rmdir(lock)
    OUT.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source(), 'cases': found}, indent=1) + '\n', encoding='utf-8')


if __name__ == '__main__':
    main()
