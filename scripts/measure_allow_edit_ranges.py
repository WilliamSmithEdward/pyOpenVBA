"""Protection.AllowEditRanges: ranges a protected sheet still lets a macro or a person edit, what each member answers, how
the sheet's protection treats them, and what a save writes.

Each case runs on a new workbook saved to a temporary folder, and its sheet part's protectedRanges element is kept where
the case saves. Users, which name this machine's accounts, are read only as a count.

    python scripts/measure_allow_edit_ranges.py

writes tests/fixtures/allow_edit_ranges.json, which tests/test_excel_allow_edit_ranges.py replays.
"""
from __future__ import annotations

import json
import re
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/allow_edit_ranges.json'
E = 'ws.Protection.AllowEditRanges'
A = f'{E}.Add "One", ws.Range("A1:B2")'
CASES = [
    f'r = TypeName({E})', f'r = {E}.Count', f'r = {E}(1).Title', f'r = {E}("x").Title', f'r = {E}.Item(0).Title', f'For Each x In {E}: i = i + 1: Next: r = i', f'r = TypeName({E}.Parent)', f'r = TypeName({E}.Application)',
    f'r = TypeName({E}.Add("One", ws.Range("A1")))', f'{A}: r = {E}.Count', f'{A}: r = {E}(1).Title', f'{A}: r = {E}("One").Title', f'{A}: r = {E}("ONE").Title', f'{A}: r = {E}.Item(1).Title', f'{A}: r = {E}(1)', f'{A}: r = TypeName({E}(1))', f'{A}: r = {E}(1).Range.Address', f'{A}: r = {E}(1).Range.Parent.Name', f'{A}: r = TypeName({E}(1).Users)', f'{A}: r = {E}(1).Users.Count', f'{A}: r = TypeName({E}(1).Parent)', f'{A}: r = TypeName({E}(1).Application)', f'{A}: For Each x In {E}: r = TypeName(x): Next',
    f'{A}: {A}', f'{A}: {E}.Add "ONE", ws.Range("C1"): r = {E}.Count', f'{E}.Add "", ws.Range("A1")', f'{E}.Add "One"', f'{E}.Add', f'{E}.Add "One", "A1"', f'{E}.Add "One", Nothing', f'{E}.Add 5, ws.Range("A1"): r = {E}(1).Title', f'{E}.Add String(255, "a"), ws.Range("A1"): r = Len({E}(1).Title)', f'{E}.Add String(256, "a"), ws.Range("A1")', f'{E}.Add "A B", ws.Range("A1"): r = {E}(1).Title', f'{E}.Add "1a", ws.Range("A1"): r = {E}(1).Title',
    f'{E}.Add "One", ws.Range("A1,C3"): r = {E}(1).Range.Address', f'{E}.Add "One", ws.Range("A1"), "pw": r = {E}(1).Title', f'{E}.Add "One", other.Range("A1")', f'{E}.Add "One", ws.Range("A:A"): r = {E}(1).Range.Address', f'{E}.Add "One", ws.Range("A1"): {E}.Add "Two", ws.Range("A1"): r = {E}.Count',
    # Changing.
    f'{A}: {E}(1).Title = "New": r = {E}(1).Title', f'{A}: {E}.Add "Two", ws.Range("C1"): {E}(1).Title = "Two"', f'{A}: {E}(1).Title = ""', f'{A}: {E}(1).Range = ws.Range("D4"): r = {E}(1).Range.Address', f'{A}: Set {E}(1).Range = ws.Range("D4"): r = {E}(1).Range.Address', f'{A}: {E}(1).Range = "D4"', f'{A}: {E}(1).ChangePassword "pw": r = {E}(1).Title', f'{A}: {E}(1).ChangePassword', f'{A}: {E}(1).Unprotect', f'{E}.Add "One", ws.Range("A1"), "pw": {E}(1).Unprotect "pw": r = {E}(1).Title', f'{E}.Add "One", ws.Range("A1"), "pw": {E}(1).Unprotect "no"',
    f'{A}: {E}(1).Delete: r = {E}.Count', f'{A}: r = {E}(1).Delete', f'{A}: {E}.Add "Two", ws.Range("C1"): {E}(1).Delete: r = {E}(1).Title', f'{A}: Set x = {E}(1): x.Delete: r = x.Title', f'{A}: {E}(1).Users.Add "Everyone", True: r = {E}(1).Users.Count & "/" & {E}(1).Users(1).Name & "/" & {E}(1).Users(1).AllowEdit', f'{A}: {E}(1).Users.Add "Nobody Here", True', f'{A}: {E}(1).Users.Add "Everyone", False: {E}(1).Users.RemoveAll: r = {E}(1).Users.Count',
    # Against protection.
    f'{A}: ws.Protect: ws.Range("A1").Value = 1: r = ws.Range("A1").Value', f'{A}: ws.Protect: ws.Range("C1").Value = 1', f'ws.Protect: {A}', f'{A}: ws.Protect: r = {E}.Count', f'{A}: ws.Protect: {E}(1).Delete', f'{A}: ws.Protect: {E}(1).Title = "New"', f'{A}: ws.Protect: r = {E}(1).Title', f'{A}: ws.Protect: ws.Unprotect: {E}(1).Delete: r = {E}.Count', f'{E}.Add "One", ws.Range("A1"), "pw": ws.Protect: ws.Range("A1").Value = 1', f'{E}.Add "One", ws.Range("A1"), "pw": ws.Protect: {E}(1).Unprotect "pw": ws.Range("A1").Value = 1: r = ws.Range("A1").Value', f'{A}: ws.Protect: r = ws.Range("A1").Locked', f'{A}: ws.Protect: ws.Range("A1").Clear', f'{A}: ws.Protect: ws.Range("A1:C1").Value = 1', f'{A}: ws.Protect: ws.Range("A1:B2").Value = 1: r = ws.Range("B2").Value', f'{A}: ws.Protect: ws.Range("A1").Interior.ColorIndex = 3',
    # Structure.
    f'{A}: ws.Rows(1).Insert: r = {E}(1).Range.Address', f'{A}: ws.Columns(1).Delete: r = {E}(1).Range.Address', f'{E}.Add "One", ws.Range("A1"): ws.Columns(1).Delete: r = {E}.Count', f'{A}: ws.Range("A1:B2").Cut ws.Range("D5"): r = {E}(1).Range.Address', f'{A}: ws.Copy After:=ws: r = ActiveSheet.Protection.AllowEditRanges.Count & "/" & ActiveSheet.Protection.AllowEditRanges(1).Title & "/" & ActiveSheet.Protection.AllowEditRanges(1).Range.Parent.Name',
    # Saves.
    f'{A}: w2.Save', f'{A}: {E}.Add "Two", ws.Range("C1,E5"): w2.Save', f'{E}.Add "One", ws.Range("A1"), "pw": w2.Save', f'{A}: ws.Protect: w2.Save', f'{A}: {E}(1).Users.Add "Everyone", True: w2.Save', f'{A}: {E}(1).Delete: w2.Save',
    f'{A}: {E}.Add "Two", ws.Range("C1"), "pw": w2.Save: n = w2.FullName: w2.Close False: Set w2 = Workbooks.Open(n): Set ws = w2.Worksheets(1): r = {E}.Count & "/" & {E}(1).Title & "/" & {E}(1).Range.Address & "/" & {E}(2).Title & "/" & {E}(2).Range.Address',
]

MODULE = r'''Option Explicit

Private Function Shown(ByVal v As Variant) As String
    On Error Resume Next
    If IsError(v) Then
        Shown = CStr(v)
    ElseIf IsArray(v) Then
        Shown = "array"
    Else
        Shown = TypeName(v) & ":" & CStr(v)
    End If
End Function

Public Function Tried(ByVal which As Long, ByVal path As String) As String
    Dim w2 As Object, ws As Object, other As Object, r As Variant, s As String
    Application.DisplayAlerts = False
    Set w2 = Workbooks.Add(xlWBATWorksheet)
    Set ws = w2.Worksheets(1)
    Set other = w2.Worksheets.Add(After:=ws)
    other.Name = "Other"
    ws.Activate
    w2.SaveAs path, xlOpenXMLWorkbook
    On Error GoTo Failed
    Run which, w2, ws, other, r
    s = "ok " & Shown(r)
Answered:
    On Error Resume Next
    Application.DisplayAlerts = False
    Tried = s
    w2.Close False
    Exit Function
Failed:
    s = "Err " & Err.Number & " " & Err.Description
    Resume Answered
End Function
'''


def _source() -> str:
    lines = ['', 'Private Sub Run(ByVal which As Long, w2 As Object, ws As Object, other As Object, r As Variant)',
             '    Dim x As Object, i As Long, n As String', '    Select Case which']
    for index, statement in enumerate(CASES):
        lines += [f'    Case {index}', f'        {statement}']
    lines += ['    End Select', 'End Sub', '']
    return MODULE + '\n'.join(lines)


def _sheet_part(path: Path) -> str:
    with zipfile.ZipFile(path) as saved:
        found = re.search(r'<protectedRanges>.*?</protectedRanges>', saved.read('xl/worksheets/sheet1.xml').decode('utf-8'), re.DOTALL)
    return found.group(0) if found else ''


def main() -> None:
    source = _source()
    answers: list[str] = []
    saved: dict[str, str] = {}
    with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as folder:
        index = 0
        while index < len(CASES):
            with ExcelSession(HarnessConfig(exclusive=False)) as excel:
                excel.new_document()
                while index < len(CASES):
                    path = Path(folder) / f'Book{index}.xlsx'
                    result = excel.run_vba(source, 'Tried', (index, str(path)), timeout=25.0)
                    index += 1
                    if not result.ok or result.dialogs:
                        text = '; '.join(getattr(dialog, 'message', '') for dialog in result.dialogs)
                        answers.append(f'breaks: {result.outcome} {text}')
                        print(index - 1, CASES[index - 1], '=>', answers[-1], flush=True)
                        break
                    answers.append(re.sub(re.escape(folder), '<folder>', str(result.value), flags=re.IGNORECASE))
                    if 'w2.Save' in CASES[index - 1]:
                        saved[str(index - 1)] = _sheet_part(path)
                    print(index - 1, CASES[index - 1], '=>', answers[-1], saved.get(str(index - 1), ''), flush=True)
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source, 'saved': saved,
              'cases': [[case, answer] for case, answer in zip(CASES, answers, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')


if __name__ == '__main__':
    main()
