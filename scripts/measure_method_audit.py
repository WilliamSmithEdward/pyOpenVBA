"""Methods held to the checklist's VERIFIED standard (docs/host_completeness.md): each with no argument and with each kind
of value for each argument, on an object of its own, late and early bound: the error, what it answers, what the object
then reads, and the same after a save and reopen.

    python scripts/measure_method_audit.py

writes tests/fixtures/method_audit.json, which tests/test_excel_method_audit.py replays. Its "early" holds the answers
that differ with o typed.
"""
from __future__ import annotations

import json
import re
import tempfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/method_audit.json'
NUMBERS = ['1', '0', '-1', '99', '1.5', '"2"', '"x"', '""', 'Null', 'Empty', 'True']
TEXTS = ['"abc"', '""', '5', 'True', 'Null', 'Empty', 'String(256, "a")']
_TABLE = ('ws.Range("B2:C2").Value = Array("x", "y"): ws.Range("B3:C3").Value = Array(1, 2): '
          'ws.Range("B4:C4").Value = Array(3, 4): ws.ListObjects.Add xlSrcRange, ws.Range("B2:C4"), , xlYes')
_NOTES = 'ws.Range("B2").AddComment "one": ws.Range("B4").AddComment "two"'
_LINK = 'ws.Hyperlinks.Add ws.Range("B2"), "http://example.com/": ws.Hyperlinks.Add ws.Range("B4"), "http://example.org/"'
_RULE = 'ws.Range("B2").Validation.Add xlValidateList, xlValidAlertStop, xlBetween, "a,b"'
_FORMATTED = 'ws.Range("B2").Value = 1: ws.Range("B2").Font.Bold = True: ws.Range("B2").Interior.Color = 255'
_LINES = 'ws.Range("B2").Value = 1: ws.Range("B4").Value = 2'
_SORT = 'ws.Range("B2:B3").Value = 1: ws.Sort.SortFields.Add ws.Range("B2")'
#: (the class, the member, its type, the object, setup, what to read after, calls). A call's ARG stands for each value of
#: its list; "r =" keeps what it answers.
METHODS = [
    ('Application', 'CentimetersToPoints', 'Application', 'Application', '', [], [('r = o.CentimetersToPoints(ARG)', NUMBERS), ('r = o.CentimetersToPoints', [])]),
    ('Application', 'InchesToPoints', 'Application', 'Application', '', [], [('r = o.InchesToPoints(ARG)', NUMBERS), ('r = o.InchesToPoints', [])]),
    ('Comment', 'Delete', 'Comment', 'ws.Range("B2").Comment', _NOTES, ['ws.Comments.Count'], [('o.Delete', []), ('o.Delete ARG', ['1'])]),
    ('Comment', 'Next', 'Comment', 'ws.Range("B2").Comment', _NOTES, [], [('r = TypeName(o.Next) & ":" & o.Next.Text', []), ('r = o.Next(ARG)', ['1'])]),
    ('Comment', 'Previous', 'Comment', 'ws.Range("B4").Comment', _NOTES, [], [('r = TypeName(o.Previous) & ":" & o.Previous.Text', []), ('r = o.Previous(ARG)', ['1'])]),
    ('Comment', 'Text', 'Comment', 'ws.Range("B2").Comment', _NOTES, ['ws.Range("B2").Comment.Text'],
     [('r = o.Text', []), ('r = o.Text(ARG)', TEXTS), ('r = o.Text("abc", ARG)', NUMBERS), ('r = o.Text("abc", 2, ARG)', ['True', 'False', 'Null', 'Empty', '"x"', '1'])]),
    ('Hyperlink', 'Delete', 'Hyperlink', 'ws.Hyperlinks(1)', _LINK, ['ws.Hyperlinks.Count'], [('o.Delete', [])]),
    ('Hyperlinks', 'Delete', 'Hyperlinks', 'ws.Hyperlinks', _LINK, ['ws.Hyperlinks.Count'], [('o.Delete', [])]),
    ('ListColumns', 'Add', 'ListColumns', 'ws.ListObjects(1).ListColumns', _TABLE, ['ws.ListObjects(1).ListColumns.Count', 'ws.ListObjects(1).Range.Address'],
     [('r = TypeName(o.Add)', []), ('r = TypeName(o.Add(ARG))', NUMBERS), ('r = TypeName(o.Add(1, ARG))', ['True', 'False', 'Null', '"x"'])]),
    ('ListRows', 'Add', 'ListRows', 'ws.ListObjects(1).ListRows', _TABLE, ['ws.ListObjects(1).ListRows.Count', 'ws.ListObjects(1).Range.Address'],
     [('r = TypeName(o.Add)', []), ('r = TypeName(o.Add(ARG))', NUMBERS), ('r = TypeName(o.Add(1, ARG))', ['True', 'False', 'Null', '"x"'])]),
    ('ListObject', 'Delete', 'ListObject', 'ws.ListObjects(1)', _TABLE, ['ws.ListObjects.Count', 'ws.Range("B2").Value', 'ws.UsedRange.Address'], [('o.Delete', [])]),
    ('Range', 'AddComment', 'Range', 'ws.Range("B2")', '', ['ws.Comments.Count', 'ws.Range("B2").Comment.Text'],
     [('r = TypeName(o.AddComment)', []), ('r = TypeName(o.AddComment(ARG))', TEXTS)]),
    ('Range', 'ClearComments', 'Range', 'ws.Range("B2:B4")', _NOTES, ['ws.Comments.Count'], [('o.ClearComments', [])]),
    ('Range', 'ClearNotes', 'Range', 'ws.Range("B2:B4")', _NOTES, ['ws.Comments.Count'], [('o.ClearNotes', [])]),
    ('Range', 'ClearFormats', 'Range', 'ws.Range("B2")', _FORMATTED, ['ws.Range("B2").Font.Bold', 'ws.Range("B2").Interior.Pattern', 'ws.Range("B2").Value'], [('o.ClearFormats', [])]),
    ('Range', 'ClearHyperlinks', 'Range', 'ws.Range("B2:B4")', _LINK, ['ws.Hyperlinks.Count'], [('o.ClearHyperlinks', [])]),
    ('Range', 'Dirty', 'Range', 'ws.Range("B2")', _LINES, ['ws.Range("B2").Value'], [('o.Dirty', [])]),
    ('Range', 'NoteText', 'Range', 'ws.Range("B2")', _NOTES, ['ws.Range("B2").Comment.Text'],
     [('r = o.NoteText', []), ('r = o.NoteText(ARG)', TEXTS), ('r = o.NoteText("abc", ARG)', NUMBERS), ('r = o.NoteText("abc", 2, ARG)', NUMBERS)]),
    ('Sort', 'SetRange', 'Sort', 'ws.Sort', _LINES, ['TypeName(ws.Sort.Rng)', 'ws.Sort.Rng.Address'],
     [('o.SetRange ws.Range("B2:B4")', []), ('o.SetRange ARG', ['Null', 'Empty', '"B2:B4"', '1', 'True', 'Application']), ('o.SetRange', [])]),
    ('SortFields', 'Clear', 'SortFields', 'ws.Sort.SortFields', _SORT, ['ws.Sort.SortFields.Count'], [('o.Clear', [])]),
    ('Validation', 'Delete', 'Validation', 'ws.Range("B2").Validation', _RULE, ['ws.Range("B2").Validation.Type'], [('o.Delete', [])]),
    ('Window', 'Activate', 'Window', 'wb.Windows(1)', 'ThisWorkbook.Activate', ['ActiveWorkbook Is wb'],
     [('o.Activate', []), ('o.Activate 1', [])]),
]


WRONG_COUNT = (r'^r = o\.(CentimetersToPoints|InchesToPoints)$|^o\.Delete 1$|^o\.Activate 1$|o\.(Next|Previous)\(1\)|^o\.SetRange( (Null|Empty|"B2:B4"|1|True|Application))?$')


def cases() -> list[tuple[str, str, str, str, str, list[str], str]]:
    """Each (class, member, type, object, setup, reads, call with its value)."""
    found = []
    for owner, member, kind, target, setup, reads, calls in METHODS:
        for call, values in calls:
            for value in values or [None]:
                found.append((owner, member, kind, target, setup, reads, call if value is None else call.replace('ARG', value)))
    return found


def _read(number: int, reads: list[str]) -> str:
    lines = ''.join(f'    v = Empty\n    v = {read}\n    If Err.Number <> 0 Then\n        s = s & "|e" & Err.Number\n'
                    f'        Err.Clear\n    Else\n        s = s & "|" & Show(v)\n    End If\n' for read in reads)
    return (f'\nPrivate Function R{number}(ws As Object) As String\n    Dim s As String, v As Variant\n'
            f'    On Error Resume Next\n{lines}    R{number} = s\nEnd Function\n')


def source(typed: bool = False) -> str:
    """Setup makes the book; S(i) makes case i on a sheet of its own; Reopen saves; A(i) reads the case again."""
    body = ('Private wb As Object\n\nPublic Sub Setup()\n    Set wb = Workbooks.Add(xlWBATWorksheet)\nEnd Sub\n'
            '\nPublic Sub Reopen(ByVal folder As String)\n    Dim path As String\n'
            '    path = folder & Application.PathSeparator & "methods.xlsx"\n'
            '    wb.SaveAs path, xlOpenXMLWorkbook\n    wb.Close False\n    Set wb = Workbooks.Open(path)\nEnd Sub\n'
            '\nPublic Sub Finish()\n    wb.Close False\nEnd Sub\n'
            '\nPrivate Function Show(v As Variant) As String\n    If IsNull(v) Then\n        Show = "Null"\n'
            '    ElseIf IsEmpty(v) Then\n        Show = "Empty"\n    ElseIf Len(CStr(v)) > 40 Then\n'
            '        Show = TypeName(v) & " " & Len(CStr(v))\n    Else\n        Show = TypeName(v) & " " & CStr(v)\n'
            '    End If\nEnd Function\n')
    for number, (_, _, kind, target, setup, reads, call) in enumerate(cases()):
        # A wrong number or type of arguments does not compile on a typed variable, so only late binding can answer it.
        statement = 'r = "does not compile"' if typed and (
            re.search(WRONG_COUNT, call) or (kind == 'ListColumns' and ', ' in call)) else call
        body += _read(number, reads)
        body += (f'\nPublic Function S{number}() As String\n    Dim ws As Object, o As {kind if typed else "Object"}, r As Variant, e As Long\n'
                 f'    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))\n'
                 + (f'    {setup}\n' if setup else '')
                 + f'    Set o = {target}\n    On Error Resume Next\n    {statement}\n    e = Err.Number\n    Err.Clear\n'
                 f'    S{number} = "e" & e & "|" & Show(r) & R{number}(ws)\nEnd Function\n'
                 f'\nPublic Function A{number}() As String\n    A{number} = R{number}(wb.Worksheets({number + 2}))\nEnd Function\n')
    return body


def _runner(calls: list[str], name: str) -> str:
    """A function answering each of the calls joined by @@, split up since one procedure cannot hold them all."""
    parts = [calls[at:at + 200] for at in range(0, len(calls), 200)]
    body = ''
    for number, part in enumerate(parts):
        lines = ''.join(f'    out = out & "@@" & {call}()\n' for call in part)
        body += (f'\nPrivate Function {name}{number}() As String\n    Dim out As String\n{lines}'
                 f'    {name}{number} = Mid(out, 3)\nEnd Function\n')
    joined = ' & "@@" & '.join(f'{name}{number}()' for number in range(len(parts)))
    return body + f'\nPrivate Function {name}() As String\n    {name} = {joined}\nEnd Function\n'


def main() -> None:
    every = cases()
    count = len(every)
    runner = _runner([f'S{number}' for number in range(count)], 'First') + _runner([f'A{number}' for number in range(count)], 'Second')
    runner += ('\nPublic Function Probe(ByVal folder As String) As String\n    Dim out As String\n    Setup\n'
               '    out = First()\n    Reopen folder\n    out = out & "@@" & Second()\n    Finish\n    Probe = out\nEnd Function\n')
    only = _runner([f'S{number}' for number in range(count)], 'First') + (
        '\nPublic Function Probe() As String\n    Setup\n    Probe = First()\n    Finish\nEnd Function\n')
    with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=1200.0, exclusive=False)) as excel:
        excel.new_document()
        result = excel.run_vba(source() + runner, 'Probe', (folder,), timeout=400.0)
        assert result.ok and not result.dialogs, result
        bound = excel.run_vba(source(typed=True) + only, 'Probe', (), timeout=400.0)
        assert bound.ok and not bound.dialogs, bound
    answers = str(result.value).split('@@')
    early = str(bound.value).split('@@')
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source(),
              'early': {str(number): early[number] for number in range(count)
                        if early[number] != answers[number] and 'does not compile' not in early[number]},
              'cases': [[every[number][0], every[number][1], every[number][6], answers[number], answers[count + number]]
                        for number in range(count)]}
    OUT.write_text(json.dumps(record, indent=0) + '\n', encoding='utf-8')
    print(count, 'cases', len(record['early']), 'differ when typed')


if __name__ == '__main__':
    main()
