"""Workbook.UserStatus's items and Workbook.ReadOnly after each way a macro reaches it, in live Excel, typed and late, with the model's answers
written down first.

UserStatus is read on a new book, after SaveAs, and opened read-write: its bounds and each item's TypeName, and whether the first item is
Application.UserName and the second within two minutes of Now, each as a yes or no, never the name or the time. ReadOnly is read on a new
book, after SaveAs, opened read-only and read-write, after SaveAs of a book opened read-only, after ChangeFileAccess both ways, and on a file
whose read-only attribute is set on disk. Every file is in a temporary folder; the attribute is cleared before the folder goes.
ChangeFileAccess back to read and write is left out: the model refuses it openly, as Excel reopens the file for it.

    python scripts/measure_workbook_status_reads.py predict
    python scripts/measure_workbook_status_reads.py measure

writes tests/fixtures/workbook_status_reads.json.
"""
from __future__ import annotations

import hashlib
import json
import shutil
import sys
import tempfile
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / 'tests/fixtures/workbook_status_reads.json'
SEPARATOR = '%%'


def macro(declared: str) -> str:
    return f'''Private Function Status(ByVal wb As {declared}) As String
    Dim v As Variant, s As String
    On Error Resume Next
    v = wb.UserStatus
    If Err.Number <> 0 Then
        Status = "e" & Err.Number
        Exit Function
    End If
    s = LBound(v, 1) & "-" & UBound(v, 1) & "x" & LBound(v, 2) & "-" & UBound(v, 2)
    s = s & "/" & TypeName(v(1, 1)) & ":" & (v(1, 1) = Application.UserName)
    s = s & "/" & TypeName(v(1, 2)) & ":" & (Abs(CDbl(CDate(v(1, 2))) - CDbl(Now)) * 86400 < 120)
    s = s & "/" & TypeName(v(1, 3)) & ":" & CStr(v(1, 3))
    Status = s
End Function
Private Function Flag(ByVal wb As {declared}) As String
    On Error Resume Next
    Dim v As Variant
    v = wb.ReadOnly
    Flag = IIf(Err.Number, "e" & Err.Number, TypeName(v) & ":" & CStr(v))
End Function
Function Probe(ByVal folder As String) As String
    Dim wb As {declared}, out As String, e As Long
    Application.DisplayAlerts = False
    Set wb = Workbooks.Add
    out = "new " & Status(wb) & " " & Flag(wb)
    wb.SaveAs folder & "\\a.xlsx", xlOpenXMLWorkbook
    out = out & "{SEPARATOR}saved " & Status(wb) & " " & Flag(wb)
    wb.Close SaveChanges:=False
    Set wb = Workbooks.Open(folder & "\\a.xlsx")
    out = out & "{SEPARATOR}read-write " & Status(wb) & " " & Flag(wb)
    On Error Resume Next
    wb.ChangeFileAccess xlReadOnly
    e = Err.Number
    On Error GoTo 0
    out = out & "{SEPARATOR}to read-only e" & e & " " & Flag(wb)
    wb.Close SaveChanges:=False
    Set wb = Workbooks.Open(folder & "\\a.xlsx", ReadOnly:=True)
    out = out & "{SEPARATOR}read-only " & Flag(wb)
    wb.SaveAs folder & "\\b.xlsx", xlOpenXMLWorkbook
    out = out & "{SEPARATOR}read-only saved as " & Flag(wb)
    wb.Close SaveChanges:=False
    SetAttr folder & "\\a.xlsx", vbReadOnly
    Set wb = Workbooks.Open(folder & "\\a.xlsx")
    out = out & "{SEPARATOR}attribute " & Flag(wb)
    wb.Close SaveChanges:=False
    SetAttr folder & "\\a.xlsx", vbNormal
    Application.DisplayAlerts = True
    Probe = out
End Function
'''


def predict() -> None:
    sys.path.insert(0, str(ROOT / 'src'))
    from pyopenvba.apps.excel import ExcelApplication

    answers: dict[str, list[str]] = {}
    for declared in ('Workbook', 'Object'):
        folder = Path(tempfile.mkdtemp(prefix='wbstatus_'))
        try:
            app = ExcelApplication(sandbox=False)
            app.add_workbook()
            app.add_module(macro(declared), name='H')
            answers[declared] = str(app.run('H.Probe', str(folder))).split(SEPARATOR)
        finally:
            shutil.rmtree(folder, ignore_errors=True)
    text = json.dumps({'predicted': answers}, indent=1) + '\n'
    OUT.write_text(text, encoding='utf-8', newline='\n')
    print('predictions sha256', hashlib.sha256(text.encode('utf-8')).hexdigest())
    print(json.dumps(answers, indent=1))


def measure() -> None:
    from pyvbaharness import ExcelSession, HarnessConfig

    record: dict[str, Any] = json.loads(OUT.read_text(encoding='utf-8'))
    assert 'predicted' in record, 'write the predictions first'
    answers: dict[str, list[str]] = {}
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        excel.run_vba('Function Hello() As String\n    Hello = "x"\nEnd Function\n', 'Hello', (), timeout=60)
        for declared in ('Workbook', 'Object'):
            folder = Path(tempfile.mkdtemp(prefix='wbstatus_'))
            try:
                result = excel.run_vba(macro(declared), 'Probe', (str(folder),), timeout=180, line_numbers=False)
                assert result.ok, (result.error, [one.message for one in result.dialogs])
                answers[declared] = str(result.value).split(SEPARATOR)
            finally:
                shutil.rmtree(folder, ignore_errors=True)
    record.update({'measured_at': datetime.now(timezone.utc).isoformat(), 'excel': answers})
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8', newline='\n')
    for declared in answers:
        for mine, theirs in zip(record['predicted'][declared], answers[declared]):
            print(declared, 'OK  ' if mine == theirs else 'MISS', ascii(theirs), '' if mine == theirs else f'model {ascii(mine)}')


if __name__ == '__main__':
    {'predict': predict, 'measure': measure}[sys.argv[1]]()
