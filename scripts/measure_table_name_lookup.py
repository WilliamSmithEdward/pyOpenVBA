"""ListObjects(text) when a table's Name and DisplayName differ, as setting Name to a cell reference ("Tb7"), to text with a space
("a b") or to text starting with a digit ("1abc") makes them. Rival rules: the lookup goes by DisplayName only, as the model had it,
which held-out round 51101 refuted once; by Name only; or by either. Each is read on a new book and again after Excel's own save
and reopen, in the case given and in upper case. Answers carry TypeName.

    python scripts/measure_table_name_lookup.py

Takes excel-coordination/live.lock. Writes tests/fixtures/table_name_lookup.json, which tests/test_excel_table_name_lookup.py replays.
"""
from __future__ import annotations

import json
import os
import subprocess
import tempfile
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / 'tests' / 'fixtures' / 'table_name_lookup.json'
NAMES = ['Tb7', 'ab12', 'a b', '1abc', 'Plain', 'R1C1', 'xfd1']


def source() -> str:
    """Probe(path, wanted, reopen) names the table and reads Name, DisplayName and each lookup."""
    return '''
Private Function Look(ByVal ws As Object, ByVal key As String) As String
    Dim t As Object
    On Error Resume Next
    Set t = ws.ListObjects(key)
    If Err.Number <> 0 Then Look = "E" & Err.Number Else Look = TypeName(t) & ":" & t.Name
End Function

Public Function Probe(ByVal path As String, ByVal wanted As String, ByVal reopen As Boolean) As String
    Dim wb As Object, ws As Object, lo As Object, shown As String, s As String
    Application.DisplayAlerts = False
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Worksheets(1)
    ws.Range("E1:F1").Value = Array("h", "g"): ws.Range("E2:F3").Value = 1
    Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("E1:F3"), , xlYes)
    On Error Resume Next
    lo.Name = wanted
    s = "set E" & Err.Number
    Err.Clear
    On Error GoTo 0
    If reopen Then
        wb.SaveAs path, xlOpenXMLWorkbook
        wb.Close False
        Set wb = Workbooks.Open(path)
        Set ws = wb.Worksheets(1)
        Set lo = ws.ListObjects(1)
    End If
    shown = lo.DisplayName
    s = s & "|" & TypeName(lo.Name) & ":" & lo.Name & "|" & TypeName(shown) & ":" & shown
    s = s & "|" & Look(ws, wanted) & "|" & Look(ws, UCase(wanted)) & "|" & Look(ws, shown) & "|" & Look(ws, UCase(shown))
    wb.Close False
    Probe = s
End Function
'''


def _lock() -> Path:
    common = subprocess.run(['git', 'rev-parse', '--git-common-dir'], capture_output=True, text=True, check=True, cwd=ROOT).stdout.strip()
    return (ROOT / common).resolve() / 'excel-coordination' / 'live.lock'


def main() -> None:
    from pyvbaharness import ExcelSession, HarnessConfig

    cases: list[dict[str, object]] = []
    lock = _lock()
    os.mkdir(lock)
    try:
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as held, ExcelSession(HarnessConfig(lock_wait_s=1200.0, exclusive=False)) as excel:
            excel.new_document()
            for at, wanted in enumerate(NAMES):
                for reopen in (False, True):
                    done = excel.run_vba(source(), 'Probe', (str(Path(held) / f'n{at}.xlsx'), wanted, reopen), timeout=60.0, line_numbers=False)
                    assert done.ok, done
                    cases.append({'name': wanted, 'reopen': reopen, 'answer': str(done.value)})
                    print(wanted, reopen, done.value, flush=True)
    finally:
        os.rmdir(lock)
    OUT.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source(), 'cases': cases}, indent=1) + '\n', encoding='utf-8')


if __name__ == '__main__':
    main()
