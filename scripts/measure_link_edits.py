"""Workbook.UpdateLink, ChangeLink, OpenLinks and SetLinkOnData on a workbook whose links' sources are closed.

Alpha.xlsx and Beta.xlsx are what Dest.xlsx reads, Gamma.xlsx has the same sheet with other values, Other.xlsx has no
sheet Dest reads, and Plain.xlsx links to nothing. Dest is saved with its sources open; Alpha is then changed and saved,
so a link Dest keeps is out of date and an update shows. Each case reopens Dest alone, its links not updated, and reads
back its formulas, values, names, links and the open workbooks.

A change to a workbook without a sheet the formulas read asks which sheet, and a missing file asks where it is. Both are
dialogs, and are recorded as what the run said. Paths are written <folder>.

    python scripts/measure_link_edits.py

writes tests/fixtures/link_edits.json, which tests/test_excel_link_edits.py replays.
"""
from __future__ import annotations

import base64
import json
import re
import tempfile
import time
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/link_edits.json'
FILES = ('Alpha.xlsx', 'Beta.xlsx', 'Gamma.xlsx', 'Other.xlsx', 'Plain.xlsx', 'Dest.xlsx', 'Lost.xlsx')
A, B, G = 'folder & "\\Alpha.xlsx"', 'folder & "\\Beta.xlsx"', 'folder & "\\Gamma.xlsx"'

CASES = [
    # Nothing done, for the rest to differ from.
    'r = 0',
    # UpdateLink.
    'd.UpdateLink', 'r = d.UpdateLink', f'd.UpdateLink {A}', f'r = d.UpdateLink({A})', f'd.UpdateLink Name:={A}, Type:=xlExcelLinks',
    f'd.UpdateLink Name:={A}, Type:=xlOLELinks', f'd.UpdateLink {B}', 'd.UpdateLink "Alpha.xlsx"', 'd.UpdateLink "ALPHA.XLSX"',
    f'd.UpdateLink UCase({A})', f'd.UpdateLink Array({A}, {B})', f'd.UpdateLink Array({A})', 'd.UpdateLink "Nope.xlsx"', 'd.UpdateLink ""',
    'd.UpdateLink folder & "\\Nope.xlsx"', 'd.UpdateLink 5', 'd.UpdateLink Null', 'd.UpdateLink Empty', 'd.UpdateLink Type:=xlExcelLinks',
    'd.UpdateLink Type:=xlOLELinks', 'd.UpdateLink Type:=7', f'd.UpdateLink {A}, 7', f'd.UpdateLink {A}, "x"', 'd.UpdateLink d.LinkSources',
    'd.UpdateLink d.LinkSources(xlExcelLinks)', f'd.BreakLink {A}, xlExcelLinks: d.UpdateLink {A}',
    f'Workbooks.Open {A}: d.UpdateLink {A}', 'p.UpdateLink', f'p.UpdateLink {A}',
    # ChangeLink.
    f'd.ChangeLink {A}, {G}', f'r = d.ChangeLink({A}, {G})', f'd.ChangeLink {A}, {G}, xlLinkTypeExcelLinks', f'd.ChangeLink {A}, {G}, xlLinkTypeOLELinks',
    f'd.ChangeLink {A}, {G}, 7', f'd.ChangeLink "Alpha.xlsx", {G}', f'd.ChangeLink {A}, "Gamma.xlsx"', f'd.ChangeLink UCase({A}), {G}',
    f'd.ChangeLink "Nope.xlsx", {G}', f'd.ChangeLink folder & "\\Nope.xlsx", {G}', f'd.ChangeLink "", {G}', f'd.ChangeLink {A}, ""', f'd.ChangeLink {A}',
    f'd.ChangeLink {A}, {A}', f'd.ChangeLink {A}, {B}', f'd.ChangeLink {B}, {A}', f'd.ChangeLink {A}, d.FullName', f'd.ChangeLink {A}, 5', f'd.ChangeLink 5, {G}',
    f'd.ChangeLink {A}, {G}: d.ChangeLink {G}, {B}', f'd.ChangeLink {A}, {G}: d.ChangeLink {B}, {A}', f'd.ChangeLink {A}, {G}: d.UpdateLink {G}',
    f'd.BreakLink {A}, xlExcelLinks: d.ChangeLink {A}, {G}', f'Workbooks.Open {G}: d.ChangeLink {A}, {G}', f'Workbooks.Open {A}: d.ChangeLink {A}, {G}',
    f'p.ChangeLink {A}, {G}', f'd.ChangeLink {A}, {G}: d.Save: d.Close False: Set d = Workbooks.Open(folder & "\\Dest.xlsx", 0)',
    # OpenLinks.
    f'd.OpenLinks {A}', f'r = d.OpenLinks({A})', f'd.OpenLinks Name:={A}, ReadOnly:=True', f'd.OpenLinks Name:={A}, ReadOnly:=False', f'd.OpenLinks {A}, , xlExcelLinks',
    f'd.OpenLinks {A}, , xlOLELinks', f'd.OpenLinks {A}, , 7', 'd.OpenLinks "Alpha.xlsx"', f'd.OpenLinks UCase({A})', f'd.OpenLinks Array({A}, {B})', 'd.OpenLinks d.LinkSources', 'd.OpenLinks "Nope.xlsx"', 'd.OpenLinks ""',
    'd.OpenLinks', 'd.OpenLinks 5', f'd.OpenLinks {G}', f'Workbooks.Open {A}: d.OpenLinks {A}', f'd.BreakLink {A}, xlExcelLinks: d.OpenLinks {A}', f'p.OpenLinks {A}',
    f'd.OpenLinks {A}: Workbooks("Dest.xlsx").Activate', f'd.OpenLinks Array({B}, {A})',
    # SetLinkOnData, for DDE and OLE links, which these are not.
    f'd.SetLinkOnData {A}, "Macro"', f'r = d.SetLinkOnData({A}, "Macro")', f'd.SetLinkOnData {A}', 'd.SetLinkOnData "Nope"', 'd.SetLinkOnData', f'd.SetLinkOnData {A}, ""', 'p.SetLinkOnData "Nope", "Macro"',
    # A file that is not there, a workbook without the sheets, and a link whose own file has gone. Each may ask.
    f'd.ChangeLink {A}, folder & "\\Missing.xlsx"', f'd.ChangeLink {A}, folder & "\\Other.xlsx"',
    'Set l = Workbooks.Open(folder & "\\Lost.xlsx", 0): r = l.LinkSources(xlExcelLinks)(1): l.UpdateLink folder & "\\Gone.xlsx"',
    'Set l = Workbooks.Open(folder & "\\Lost.xlsx", 0): r = l.LinkSources(xlExcelLinks)(1): l.OpenLinks folder & "\\Gone.xlsx"',
    f'Set l = Workbooks.Open(folder & "\\Lost.xlsx", 0): r = l.LinkSources(xlExcelLinks)(1): l.ChangeLink folder & "\\Gone.xlsx", {G}: r = l.Worksheets(1).Range("A1").Value',
    # An edit to a source once it is open, and whether the list stays sorted after a save.
    f'd.OpenLinks {A}: Workbooks("Alpha.xlsx").Worksheets(1).Range("A1").Value = 5000', f'd.ChangeLink {A}, {G}: d.SaveCopyAs folder & "\\Copy.xlsx"',
    f'Workbooks.Open {A}: Set l = Workbooks("Alpha.xlsx"): r = l.Close(False)',
    # What a save writes, each kept below.
    f'd.UpdateLink {A}: d.Save', f'd.ChangeLink {A}, {B}: d.Save', f'd.ChangeLink {A}, folder & "\\Other.xlsx": d.Save', f'd.ChangeLink {A}, folder & "\\Missing.xlsx": d.Save',
    f'd.OpenLinks {A}: d.Save', f'd.ChangeLink {A}, "C:\\Windows\\Gamma.xlsx": d.Save', 'p.UpdateLink Type:=xlExcelLinks', f'p.OpenLinks {G}', f'p.OpenLinks {G}, , xlExcelLinks',
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

Private Function Listed(ByVal v As Variant) As String
    Dim i As Long, s As String
    If IsEmpty(v) Then
        Listed = "none"
        Exit Function
    End If
    For i = LBound(v) To UBound(v)
        s = s & "," & v(i)
    Next
    Listed = LBound(v) & ":" & Mid(s, 2)
End Function

Private Function State(ByVal d As Object) As String
    Dim s As String, cell As Range, entry As Name, b As Workbook
    On Error Resume Next
    For Each cell In d.Worksheets(1).Range("A1:A6").Cells
        s = s & "|" & cell.Formula & "=" & Shown(cell.Value)
    Next
    For Each entry In d.Names
        s = s & "|" & entry.Name & "->" & entry.RefersTo
    Next
    s = s & "|links " & Listed(d.LinkSources(xlExcelLinks))
    For Each b In Workbooks
        If b.Name <> ThisWorkbook.Name Then s = s & "|open " & b.Name & IIf(b.ReadOnly, " ro", "")
    Next
    s = s & "|active " & IIf(ActiveWorkbook.Name = ThisWorkbook.Name, "harness", ActiveWorkbook.Name)
    State = s
End Function

Private Sub Shut()
    Dim b As Workbook, again As Boolean
    On Error Resume Next
    Do
        again = False
        For Each b In Workbooks
            If b.Name <> ThisWorkbook.Name Then
                b.Close False
                again = True
                Exit For
            End If
        Next
    Loop While again
End Sub

Public Function Tried(ByVal which As Long, ByVal folder As String) As String
    Dim d As Object, p As Object, r As Variant, s As String
    Application.DisplayAlerts = False
    Application.AskToUpdateLinks = False
    Set p = Workbooks.Open(folder & "\Plain.xlsx", 0)
    Set d = Workbooks.Open(folder & "\Dest.xlsx", 0)
    Dim links As Variant
    links = d.LinkSources(xlExcelLinks)
    On Error GoTo Failed
    Run which, d, p, folder, r
    s = "ok " & Shown(r)
Answered:
    On Error GoTo 0
    Tried = s & State(d)
    Shut
    Exit Function
Failed:
    s = "Err " & Err.Number & " " & Err.Description
    Resume Answered
End Function
'''

SETUP = r'''
Private Sub Saved(ByVal b As Workbook, ByVal path As String)
    b.SaveAs path, xlOpenXMLWorkbook
    b.Close False
End Sub

Private Function Sheet(ByVal folder As String, ByVal name As String, ByVal sheetName As String, ByVal first As Double) As Workbook
    Dim b As Workbook, i As Long
    Set b = Workbooks.Add(xlWBATWorksheet)
    b.Worksheets(1).Name = sheetName
    For i = 1 To 3
        b.Worksheets(1).Cells(i, 1).Value = first * i
    Next
    b.Names.Add Name:="Rate", RefersTo:="=" & sheetName & "!$A$1"
    b.SaveAs folder & "\" & name, xlOpenXMLWorkbook
    Set Sheet = b
End Function

Private Sub LostBook(ByVal folder As String)
    Dim gone As Workbook, l As Workbook
    Set gone = Sheet(folder, "Gone.xlsx", "Sheet1", 2)
    Set l = Workbooks.Add(xlWBATWorksheet)
    l.Worksheets(1).Range("A1").Formula = "=[Gone.xlsx]Sheet1!A1"
    Saved l, folder & "\Lost.xlsx"
    gone.Close False
    Kill folder & "\Gone.xlsx"
End Sub

Public Function Setup(ByVal folder As String) As String
    Dim alpha As Workbook, beta As Workbook, d As Workbook, ws As Worksheet, p As Workbook
    Application.DisplayAlerts = False
    Set alpha = Sheet(folder, "Alpha.xlsx", "Sheet1", 1)
    Set beta = Sheet(folder, "Beta.xlsx", "Sheet1", 10)
    Sheet(folder, "Gamma.xlsx", "Sheet1", 7).Close False
    Sheet(folder, "Other.xlsx", "Elsewhere", 3).Close False
    Set p = Workbooks.Add(xlWBATWorksheet)
    p.Worksheets(1).Range("A1").Value = 1
    Saved p, folder & "\Plain.xlsx"
    Set d = Workbooks.Add(xlWBATWorksheet)
    Set ws = d.Worksheets(1)
    ws.Name = "Dest"
    d.Names.Add Name:="Ext", RefersTo:="=[Alpha.xlsx]Sheet1!$A$1:$A$3"
    ws.Range("A1").Formula = "=[Alpha.xlsx]Sheet1!A1"
    ws.Range("A2").Formula = "=[Beta.xlsx]Sheet1!A1"
    ws.Range("A3").Formula = "=[Alpha.xlsx]Sheet1!A2+[Beta.xlsx]Sheet1!A2"
    ws.Range("A4").Formula = "=SUM(Ext)"
    ws.Range("A5").Formula = "=[Alpha.xlsx]Sheet1!A3*2"
    ws.Range("A6").Value = 5
    Saved d, folder & "\Dest.xlsx"
    LostBook folder
    beta.Close False
    alpha.Worksheets(1).Range("A1").Value = 100
    alpha.Worksheets(1).Range("A2").Value = 200
    alpha.Save
    alpha.Close False
    Setup = "done"
End Function
'''


def _source() -> str:
    lines = ['', 'Private Sub Run(ByVal which As Long, d As Object, p As Object, ByVal folder As String, r As Variant)',
             '    Dim l As Object', '    Select Case which']
    for index, statement in enumerate(CASES):
        lines += [f'    Case {index}', f'        {statement}']
    lines += ['    End Select', 'End Sub', '']
    return MODULE + '\n'.join(lines) + SETUP


def _clean(text: object, folder: str) -> str:
    return re.sub(re.escape(folder), '<folder>', str(text), flags=re.IGNORECASE)


def _link_parts(path: Path, folder: str) -> dict[str, str]:
    """A saved Dest's parts that name its links, the folder written <folder>."""
    with zipfile.ZipFile(path) as saved:
        return {name: _clean(saved.read(name).decode('utf-8'), folder).replace(_clean(folder.replace('\\', '/'), folder), '<folder>')
                for name in saved.namelist() if 'externalLink' in name or name in ('xl/workbook.xml', 'xl/_rels/workbook.xml.rels')}


def _restored(path: Path, data: bytes) -> None:
    """Dest as it was saved, for the case that saves over it. Excel lets go of a closed file a moment later."""
    for _ in range(50):
        try:
            path.write_bytes(data)
            return
        except PermissionError:
            time.sleep(0.2)
    path.write_bytes(data)


def main() -> None:
    """Each case alone, in a session that a dialog or a broken run ends."""
    source = _source()
    answers: list[str] = []
    saved: dict[str, dict[str, str]] = {}
    with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as temporary:
        folder = temporary
        config = HarnessConfig(exclusive=False)
        with ExcelSession(config) as excel:
            excel.new_document()
            done = excel.run_vba(source, 'Setup', (folder,), timeout=120.0)
            if not done.ok:
                raise SystemExit(f'setup failed: {done.outcome} {done.error}')
        files = {name: base64.b64encode((Path(folder) / name).read_bytes()).decode('ascii') for name in FILES}
        dest = (Path(folder) / 'Dest.xlsx').read_bytes()
        index = 0
        while index < len(CASES):
            with ExcelSession(config) as excel:
                excel.new_document()
                while index < len(CASES):
                    _restored(Path(folder) / 'Dest.xlsx', dest)
                    result = excel.run_vba(source, 'Tried', (index, folder), timeout=25.0)
                    index += 1
                    if not result.ok or result.dialogs:
                        answers.append(f'breaks: {result.outcome} {_clean(result.dialogs, folder)}')
                        print(index - 1, CASES[index - 1], '=>', answers[-1], flush=True)
                        break
                    answers.append(_clean(result.value, folder))
                    if 'd.Save' in CASES[index - 1]:
                        saved[str(index - 1)] = _link_parts(Path(folder) / 'Dest.xlsx', folder)
                    print(index - 1, CASES[index - 1], '=>', answers[-1], flush=True)
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source, 'files': files, 'saved': saved,
              'cases': [[case, answer] for case, answer in zip(CASES, answers, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')


if __name__ == '__main__':
    main()
