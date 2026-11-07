"""FormatCondition.Parent when the range a macro reaches the rule through is not the range the rule applies to. Rival rules: Parent
is AppliesTo, as the model had it, which held-out round 51102 refuted once; it is the range FormatConditions was taken from; or
it is where the two meet. A rule on D1:D5 is reached through D1:D5, D3, C1:E9, D4:D8 and the whole sheet, through the object
Add gives back, and through D1:D5 after a row is inserted above it, each on a new book and again after Excel's own save and
reopen. Answers carry TypeName.

    python scripts/measure_format_condition_parent.py

Takes excel-coordination/live.lock. Writes tests/fixtures/format_condition_parent.json, which
tests/test_excel_format_condition_parent.py replays.
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
OUT = ROOT / 'tests' / 'fixtures' / 'format_condition_parent.json'
#: How the rule is reached, after its sheet is ready.
REACHES = ['ws.Range("D1:D5").FormatConditions(1)', 'ws.Range("D3").FormatConditions(1)', 'ws.Range("C1:E9").FormatConditions(1)',
           'ws.Range("D4:D8").FormatConditions(1)', 'ws.Cells.FormatConditions(1)', 'made',
           'inserted', 'top10 through D3', 'databar through D3', 'colour scale through C1:E9']


def source() -> str:
    return '''
Private Function Shown(ByVal v As Object) As String
    On Error Resume Next
    Shown = TypeName(v) & ":" & v.Address(False, False)
    If Err.Number <> 0 Then Shown = TypeName(v) & ":E" & Err.Number
End Function

Public Function Probe(ByVal path As String, ByVal which As Long, ByVal reopen As Boolean) As String
    Dim wb As Object, ws As Object, fc As Object, made As Object
    Application.DisplayAlerts = False
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Worksheets(1)
    ws.Range("D1:D5").Value = 3
    Select Case which
    Case 7: Set made = ws.Range("D1:D5").FormatConditions.AddTop10
    Case 8: Set made = ws.Range("D1:D5").FormatConditions.AddDatabar
    Case 9: Set made = ws.Range("D1:D5").FormatConditions.AddColorScale(3)
    Case Else: Set made = ws.Range("D1:D5").FormatConditions.Add(xlCellValue, xlGreater, "=2")
    End Select
    If reopen Then
        wb.SaveAs path, xlOpenXMLWorkbook
        wb.Close False
        Set wb = Workbooks.Open(path)
        Set ws = wb.Worksheets(1)
        Set made = Nothing
    End If
    On Error Resume Next
    Select Case which
    Case 5
        If made Is Nothing Then Probe = "none": wb.Close False: Exit Function
        Set fc = made
    Case 7, 8
        Set fc = ws.Range("D3").FormatConditions(1)
    Case 9
        Set fc = ws.Range("C1:E9").FormatConditions(1)
    Case 6
        ws.Rows(1).Insert
        Set fc = ws.Range("D1:D5").FormatConditions(1)
    Case Else
        Select Case which
''' + ''.join(f'        Case {at}: Set fc = {reach}\n' for at, reach in enumerate(REACHES[:5])) + '''        End Select
    End Select
    If Err.Number <> 0 Then Probe = "reach E" & Err.Number: Err.Clear: wb.Close False: Exit Function
    Probe = Shown(fc.Parent) & "|" & Shown(fc.AppliesTo) & "|E" & Err.Number
    wb.Close False
End Function
'''


def _lock() -> Path:
    common = subprocess.run(['git', 'rev-parse', '--git-common-dir'], capture_output=True, text=True, check=True, cwd=ROOT).stdout.strip()
    return (ROOT / common).resolve() / 'excel-coordination' / 'live.lock'


def main() -> None:
    from pyvbaharness import ExcelSession, HarnessConfig

    cases: list[dict[str, object]] = []
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
            for at, how in enumerate(REACHES):
                for reopen in (False, True):
                    done = excel.run_vba(source(), 'Probe', (str(Path(held) / f'f{at}.xlsx'), at, reopen), timeout=60.0, line_numbers=False)
                    assert done.ok, done
                    cases.append({'how': how, 'reopen': reopen, 'answer': str(done.value)})
                    print(how, reopen, done.value, flush=True)
    finally:
        os.rmdir(lock)
    OUT.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source(), 'cases': cases}, indent=1) + '\n', encoding='utf-8')


if __name__ == '__main__':
    main()
