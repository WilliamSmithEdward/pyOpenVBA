"""A formula naming a workbook the workbook has no link to yet: what it reads, what Range.Formula shows, what LinkSources
lists, what a save writes, and what the formulas read once their sources close and the workbook is opened again.

Alpha.xlsx has Sheet1 with 1, 2 and 3 in A1:A3, text, True and #N/A in B1:B3, and a name Rate for A1; Other.xlsx has a
sheet Elsewhere; Sub\\Deep.xlsx is in a folder below. Each case adds a new workbook, sets formulas on its first sheet,
some with their source open, saves it as Host.xlsx, closes every source, then closes Host and opens it again without
updating its links. Paths are written <folder>.

    python scripts/measure_book_refs.py
    python scripts/measure_book_refs.py model

writes tests/fixtures/book_refs.json, which tests/test_excel_book_refs.py replays; the second adds what Excel reads of
the model's save of each case.
"""
from __future__ import annotations

import base64
import json
import re
import sys
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/book_refs.json'
FILES = ('Alpha.xlsx', 'Other.xlsx', 'Sub/Deep.xlsx')
FULL = '"=\'" & folder & "\\[Alpha.xlsx]Sheet1\'!'

CASES = [
    # The source open, named by its file.
    'Opened folder, "Alpha.xlsx": s.Range("A1").Formula = "=[Alpha.xlsx]Sheet1!A1"',
    'Opened folder, "Alpha.xlsx": s.Range("A1").Formula = ' + FULL + 'A1"',
    'Opened folder, "Alpha.xlsx": s.Range("A1").Formula = "=SUM([Alpha.xlsx]Sheet1!A1:A3)"',
    'Opened folder, "Alpha.xlsx": s.Range("A1").Formula = "=[Alpha.xlsx]Sheet1!A1": s.Range("A2").Formula = "=[Alpha.xlsx]Sheet1!A3*2"',
    'Opened folder, "Alpha.xlsx": s.Range("A1").Formula = "=[Alpha.xlsx]Sheet1!B1": s.Range("A2").Formula = "=[Alpha.xlsx]Sheet1!B2": s.Range("A3").Formula = "=[Alpha.xlsx]Sheet1!B3"',
    'Opened folder, "Alpha.xlsx": s.Range("A1").Formula = "=[Alpha.xlsx]Sheet1!C9"',
    'Opened folder, "Alpha.xlsx": s.Range("A1").Formula = "=Alpha.xlsx!Rate"',
    'Opened folder, "Alpha.xlsx": s.Range("A1").Formula = "=[Alpha.xlsx]Nope!A1"',
    'Opened folder, "Alpha.xlsx": s.Range("A1").FormulaR1C1 = "=[Alpha.xlsx]Sheet1!R2C1"',
    'Opened folder, "Alpha.xlsx": Opened folder, "Other.xlsx": s.Range("A1").Formula = "=[Other.xlsx]Elsewhere!A1": s.Range("A2").Formula = "=[Alpha.xlsx]Sheet1!A1"',
    'Opened folder, "Alpha.xlsx": s.Range("A1").Formula = "=[Alpha.xlsx]Sheet1!A1": Workbooks("Alpha.xlsx").Worksheets(1).Range("A1").Value = 50',
    'Set n = Workbooks.Add(xlWBATWorksheet): n.Worksheets(1).Range("A1").Value = 5: s.Range("A1").Formula = "=[" & n.Name & "]Sheet1!A1"',
    # The source closed.
    's.Range("A1").Formula = ' + FULL + 'A1"',
    's.Range("A1").Formula = "=SUM(\'" & folder & "\\[Alpha.xlsx]Sheet1\'!A1:A3)"',
    's.Range("A1").Formula = ' + FULL + 'B1": s.Range("A2").Formula = ' + FULL + 'B2": s.Range("A3").Formula = ' + FULL + 'B3"',
    's.Range("A1").Formula = "=\'" & folder & "\\Sub\\[Deep.xlsx]Sheet1\'!A1"',
    's.Range("A1").Formula = "=\'" & folder & "\\Alpha.xlsx\'!Rate"',
    's.Range("A1").Formula = "=\'" & folder & "\\[Alpha.xlsx]Nope\'!A1"',
    's.Range("A1").Formula = "=\'" & folder & "\\[Nope.xlsx]Sheet1\'!A1"',
    'ChDir folder: s.Range("A1").Formula = "=[Alpha.xlsx]Sheet1!A1"',
    'ChDir folder & "\\Sub": s.Range("A1").Formula = "=[Alpha.xlsx]Sheet1!A1"',
    's.Range("A1").Formula = ' + FULL + 'A1": hostAt = "\\Sub"',
    's.Range("A1").Formula = ' + FULL + 'A1": s.Range("A1").ClearContents',
    # Which number a second link takes: the order the formulas came in, or the order of the files' names.
    'Opened folder, "Alpha.xlsx": Opened folder, "Other.xlsx": s.Range("A1").Formula = "=[Alpha.xlsx]Sheet1!A1": s.Range("A2").Formula = "=[Other.xlsx]Elsewhere!A1"',
    'Opened folder, "Alpha.xlsx": s.Range("A1").Value = "=[Alpha.xlsx]Sheet1!A2"',
    'ChDir folder & "\\Sub": s.Range("A1").Formula = "=[Deep.xlsx]Sheet1!A1"',
]

MODULE = r'''Option Explicit

Private hostAt As String

Private Function Shown(ByVal v As Variant) As String
    On Error Resume Next
    If IsError(v) Then
        Shown = CStr(v)
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
    Listed = Mid(s, 2)
End Function

Private Function State(ByVal h As Workbook) As String
    Dim s As String, cell As Range
    On Error Resume Next
    For Each cell In h.Worksheets(1).Range("A1:A3").Cells
        s = s & "|" & cell.Formula & "=" & Shown(cell.Value)
    Next
    s = s & "|links " & Listed(h.LinkSources(xlExcelLinks))
    State = s
End Function

Private Sub Opened(ByVal folder As String, ByVal name As String)
    Workbooks.Open folder & "\" & name, 0
End Sub

Private Sub Shut(ByVal keep As Workbook)
    Dim b As Workbook, again As Boolean
    On Error Resume Next
    Do
        again = False
        For Each b In Workbooks
            If b.Name <> ThisWorkbook.Name And Not b Is keep Then
                b.Close False
                again = True
                Exit For
            End If
        Next
    Loop While again
End Sub

Public Function Tried(ByVal which As Long, ByVal folder As String) As String
    Dim h As Workbook, s As Worksheet, r As String, here As String
    Application.DisplayAlerts = False
    Application.AskToUpdateLinks = False
    hostAt = ""
    here = CurDir
    Set h = Workbooks.Add(xlWBATWorksheet)
    Set s = h.Worksheets(1)
    On Error GoTo Failed
    Run which, s, folder
    r = "ok"
Answered:
    On Error GoTo 0
    r = r & State(h)
    On Error GoTo Late
    h.SaveAs folder & hostAt & "\Host.xlsx", xlOpenXMLWorkbook
    Shut h
    r = r & "||closed" & State(h)
    h.Close False
    Set h = Workbooks.Open(folder & hostAt & "\Host.xlsx", 0)
    r = r & "||reopened" & State(h)
    h.Close False
Finished:
    On Error Resume Next
    Shut Nothing
    ChDir here
    Tried = r
    Exit Function
Failed:
    r = "Err " & Err.Number & " " & Err.Description
    Resume Answered
Late:
    r = r & "||Err " & Err.Number & " " & Err.Description
    Resume Finished
End Function

Public Function ReadBack(ByVal path As String) As String
    Dim h As Workbook
    Application.DisplayAlerts = False
    Application.AskToUpdateLinks = False
    Set h = Workbooks.Open(path, 0)
    ReadBack = "reopened" & State(h)
    h.Close False
End Function

Public Function Setup(ByVal folder As String) As String
    Dim b As Workbook
    Application.DisplayAlerts = False
    Set b = Workbooks.Add(xlWBATWorksheet)
    With b.Worksheets(1)
        .Range("A1").Value = 1
        .Range("A2").Value = 2
        .Range("A3").Value = 3
        .Range("B1").Value = "word"
        .Range("B2").Value = True
        .Range("B3").Formula = "=NA()"
    End With
    b.Names.Add Name:="Rate", RefersTo:="=Sheet1!$A$1"
    b.SaveAs folder & "\Alpha.xlsx", xlOpenXMLWorkbook
    b.Close False
    Set b = Workbooks.Add(xlWBATWorksheet)
    b.Worksheets(1).Name = "Elsewhere"
    b.Worksheets(1).Range("A1").Value = 7
    b.SaveAs folder & "\Other.xlsx", xlOpenXMLWorkbook
    b.Close False
    MkDir folder & "\Sub"
    Set b = Workbooks.Add(xlWBATWorksheet)
    b.Worksheets(1).Range("A1").Value = 11
    b.SaveAs folder & "\Sub\Deep.xlsx", xlOpenXMLWorkbook
    b.Close False
    Setup = "done"
End Function
'''


def source() -> str:
    lines = ['', 'Private Sub Run(ByVal which As Long, s As Worksheet, ByVal folder As String)', '    Dim n As Workbook', '    Select Case which']
    for index, statement in enumerate(CASES):
        lines += [f'    Case {index}', f'        {statement}']
    lines += ['    End Select', 'End Sub', '']
    return MODULE + '\n'.join(lines)


def _clean(text: object, folder: str) -> str:
    """The folder written <folder>, as Windows spells it and with / for \\; a link to another folder writes it with no
    drive, /Users/.../Alpha.xlsx, which is written /<folder> (its drive is <folder>'s)."""
    text = re.sub(re.escape(folder), '<folder>', str(text), flags=re.IGNORECASE)
    text = re.sub(re.escape(folder.replace('\\', '/')), '<folder>', text, flags=re.IGNORECASE)
    return re.sub(re.escape(folder[2:].replace('\\', '/')), '/<folder>', text, flags=re.IGNORECASE)


def _link_parts(path: Path, folder: str) -> dict[str, str]:
    """The saved Host's parts that name its links, the folder written <folder>."""
    with zipfile.ZipFile(path) as saved:
        return {name: _clean(saved.read(name).decode('utf-8'), folder)
                for name in saved.namelist()
                if 'externalLink' in name or name in ('xl/workbook.xml', 'xl/_rels/workbook.xml.rels', '[Content_Types].xml', 'xl/worksheets/sheet1.xml')}


def model_saves() -> None:
    """Each case run in the model, its Host.xlsx opened in Excel beside the same files: what Excel reads of the model's
    save, kept in the fixture as excel_reads_model."""
    from pyvbaharness import ExcelSession, HarnessConfig

    from pyopenvba.apps.excel import ExcelApplication

    record = json.loads(OUT.read_text(encoding='utf-8'))
    reads: dict[str, str] = {}
    with ExcelSession(HarnessConfig(exclusive=False)) as excel:
        excel.new_document()
        for index in range(len(record['cases'])):
            with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as folder:
                for name, data in record['files'].items():
                    (Path(folder) / name).parent.mkdir(parents=True, exist_ok=True)
                    (Path(folder) / name).write_bytes(base64.b64decode(data))
                app = ExcelApplication()
                app.add_workbook()
                app.interpreter.sandbox = False
                app.add_module(record['source'], name='Probe')
                try:
                    app.run('Probe.Tried', index, folder)
                except Exception as exc:  # noqa: BLE001
                    reads[str(index)] = f'model raised {type(exc).__name__}'
                    continue
                host = next((one for one in (Path(folder) / 'Host.xlsx', Path(folder) / 'Sub/Host.xlsx') if one.is_file()), None)
                if host is None:
                    continue
                result = excel.run_vba(source(), 'ReadBack', (str(host),), timeout=30.0, line_numbers=False)
                reads[str(index)] = _clean(result.value, folder) if result.ok and not result.dialogs \
                    else f'breaks: {result.outcome} {_clean(result.dialogs, folder)}'
                print(index, '=>', reads[str(index)], flush=True)
    record['excel_reads_model'] = reads
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')


def main() -> None:
    """Each case alone, in a session that a dialog or a broken run ends."""
    from pyvbaharness import ExcelSession, HarnessConfig

    text = source()
    answers: list[str] = []
    saved: dict[str, dict[str, str]] = {}
    with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as temporary:
        folder = temporary
        config = HarnessConfig(exclusive=False)
        with ExcelSession(config) as excel:
            excel.new_document()
            done = excel.run_vba(text, 'Setup', (folder,), timeout=120.0, line_numbers=False)
            if not done.ok:
                raise SystemExit(f'setup failed: {done.outcome} {done.error}')
        files = {name: base64.b64encode((Path(folder) / name).read_bytes()).decode('ascii') for name in FILES}
        index = 0
        while index < len(CASES):
            with ExcelSession(config) as excel:
                excel.new_document()
                while index < len(CASES):
                    for host in (Path(folder) / 'Host.xlsx', Path(folder) / 'Sub/Host.xlsx'):
                        host.unlink(missing_ok=True)
                    result = excel.run_vba(text, 'Tried', (index, folder), timeout=30.0, line_numbers=False)
                    index += 1
                    if not result.ok or result.dialogs:
                        answers.append(f'breaks: {result.outcome} {_clean(result.dialogs, folder)}')
                        print(index - 1, '=>', answers[-1], flush=True)
                        break
                    answers.append(_clean(result.value, folder))
                    host = next((one for one in (Path(folder) / 'Host.xlsx', Path(folder) / 'Sub/Host.xlsx') if one.is_file()), None)
                    if host is not None:
                        saved[str(index - 1)] = _link_parts(host, folder)
                    print(index - 1, '=>', answers[-1], flush=True)
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': text, 'files': files, 'saved': saved,
              'cases': [[case, answer] for case, answer in zip(CASES, answers, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')


if __name__ == '__main__':
    if sys.argv[1:] == ["model"]:
        model_saves()
    else:
        main()
