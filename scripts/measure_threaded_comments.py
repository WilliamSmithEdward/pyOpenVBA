"""Threaded comments: what Range.AddCommentThreaded makes, what CommentThreaded, its replies and its author read,
the errors, and the parts Excel saves.

Each case makes a workbook, runs the case's statements, reads the sheet's threaded comments, and saves it. Before
anything is recorded, the signed-in user's display name and user id, the machine's user name, every GUID and every
time are replaced by placeholders: each GUID by {G1}, {G2}... in the order it first appears in the case, so the parts
still say which comment answers which. No identity of the person measuring reaches the fixture.

    python scripts/measure_threaded_comments.py

writes tests/fixtures/threaded_comments.json, which tests/test_excel_threaded_comments.py replays.
"""
from __future__ import annotations

import json
import re
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/threaded_comments.json'
CASES = [
    'ws.Range("B2").AddCommentThreaded "hello"',
    'Set c = ws.Range("B2").AddCommentThreaded("hello"): c.AddReply "one": c.AddReply "two"',
    'ws.Range("D4").AddCommentThreaded "later": ws.Range("B2").AddCommentThreaded "earlier"',
    'Set c = ws.Range("B2").AddCommentThreaded("hello"): c.Resolved = True',
    'Set c = ws.Range("B2").AddCommentThreaded("hello"): c.Text "changed"',
    'Set c = ws.Range("B2").AddCommentThreaded("hello"): c.AddReply "one": c.AddReply "two": c.Replies(1).Delete',
    'ws.Range("B2").AddCommentThreaded "hello": ws.Range("D4").AddCommentThreaded "other": '
    'ws.Range("B2").CommentThreaded.Delete',
    'ws.Range("A1").AddComment "note": ws.Range("B2").AddCommentThreaded "hello"',
    'ws.Range("B2").AddCommentThreaded "hello": ws.Range("A1").AddComment "note"',
    'ws.Range("B2").AddCommentThreaded "line1" & vbLf & "line2"',
    'ws.Range("B2").AddCommentThreaded "hello": ws.Range("B2").AddComment "note"',
    'ws.Range("B2").AddComment "note": ws.Range("B2").AddCommentThreaded "hello"',
    'ws.Range("B2").AddCommentThreaded "hello": ws.Range("B2").AddCommentThreaded "again"',
    'ws.Range("B2:C3").AddCommentThreaded "hello"',
    'ws.Range("B2").AddCommentThreaded ""',
    'Set c = ws.Range("B2").AddCommentThreaded("hello"): c.Resolved = True: c.Resolved = False',
    'Set c = ws.Range("B2").AddCommentThreaded("hello"): c.AddReply("one").Delete',
]
READS = ['ws.CommentsThreaded.Count', 'ws.Comments.Count']


def _probe() -> str:
    lines = ['Private Function Thread(ByVal c As Object) As String',
             '    Dim s As String, i As Long',
             '    s = c.Parent.Address(0, 0) & ":" & c.Text & ",resolved=" & c.Resolved & ",replies=" & c.Replies.Count',
             '    s = s & ",author=" & (c.Author.Name = who) & "/" & c.Author.ProviderID & ",date=" & TypeName(c.Date)',
             '    For i = 1 To c.Replies.Count',
             '        s = s & ";reply " & c.Replies(i).Text & "/" & TypeName(c.Replies(i).Parent)',
             '    Next', '    Thread = s', 'End Function', '',
             'Private Function Tried(ByVal which As Long, ByVal folder As String) As String',
             '    Dim wb As Workbook, ws As Worksheet, c As Object, s As String, i As Long',
             '    Set wb = Workbooks.Add(xlWBATWorksheet): Set ws = wb.Worksheets(1)',
             '    On Error GoTo Failed', '    Select Case which']
    for index, case in enumerate(CASES):
        lines += [f'    Case {index}', f'        {case}']
    lines += ['    End Select', '    On Error GoTo 0']
    lines += [f'    s = s & ",{read}=" & {read}' for read in READS]
    lines += ['    For i = 1 To ws.CommentsThreaded.Count: s = s & "|" & Thread(ws.CommentsThreaded(i)): Next',
              '    Tried = "ok" & s', 'Saved:',
              '    wb.SaveAs folder & "\\c" & which & ".xlsx", xlOpenXMLWorkbook', '    wb.Close False',
              '    Exit Function', 'Failed:', '    Tried = "Err " & Err.Number', '    Resume Saved', 'End Function', '',
              'Public Function Probe(ByVal folder As String, ByVal name As String) As String',
              '    Dim i As Long, s As String', '    who = name', '    Application.DisplayAlerts = False',
              f'    For i = 0 To {len(CASES) - 1}', '        s = s & "@@@" & Tried(i, folder)', '    Next',
              '    Application.DisplayAlerts = True', '    Probe = Mid(s, 4)', 'End Function']
    return 'Private who As String\n\n' + '\n'.join(lines) + '\n'


#: The parts a threaded comment touches, and the relationship and content type lists that name them.
PARTS = r'threadedComment|person|comments\d|vmlDrawing|\[Content_Types\]|sheet1\.xml\.rels|workbook\.xml\.rels'


def masked(text: str, secrets: list[str], guids: dict[str, str]) -> str:
    """The text with each secret, GUID and time replaced by its placeholder."""
    for secret in secrets:
        text = text.replace(secret, '<user>')
    text = re.sub(r'\d{4}-\d\d-\d\dT[\d:.]+Z?', '<time>', text)

    def guid(found: re.Match[str]) -> str:
        return guids.setdefault(found.group(0).upper(), f'G{len(guids) + 1}')

    return re.sub(r'[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}', guid, text)


def parts(path: Path, secrets: list[str]) -> dict[str, str]:
    """The touched parts of a saved workbook, masked, the person's user id among the secrets."""
    guids: dict[str, str] = {}
    found: dict[str, str] = {}
    with zipfile.ZipFile(path) as book:
        names = sorted(name for name in book.namelist() if re.search(PARTS, name))
        texts = {name: book.read(name).decode('utf-8') for name in names}
    person = next((text for name, text in texts.items() if 'person' in name), '')
    ids = re.findall(r'(?:userId|displayName)="([^"]*)"', person)
    every = sorted({*secrets, *[one for one in ids if one]}, key=len, reverse=True)
    for name in names:
        found[name] = masked(texts[name], every, guids)
    return found


def main() -> None:
    source = _probe()
    with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        user = str(excel.run_vba('Public Function U() As String\n    U = Application.UserName\nEnd Function\n',
                                 'U').value)
        # The signed-in display name, which Author.Name reads, from a first comment's author.
        named = excel.run_vba('Public Function N() As String\n    Dim wb As Workbook\n'
                              '    Set wb = Workbooks.Add(xlWBATWorksheet)\n'
                              '    N = wb.Worksheets(1).Range("A1").AddCommentThreaded("x").Author.Name\n'
                              '    wb.Close False\nEnd Function\n', 'N')
        display = str(named.value)
        result = excel.run_vba(source, 'Probe', (folder, display), timeout=280.0)
        assert result.ok and not result.dialogs, result
        answers = str(result.value).split('@@@')
        secrets = [one for one in (display, user) if one]
        saved = [parts(Path(folder) / f'c{index}.xlsx', secrets) for index in range(len(CASES))]
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source,
              'cases': [[case, masked(answer, secrets, {}), part]
                        for case, answer, part in zip(CASES, answers, saved, strict=True)]}
    text = json.dumps(record, indent=1) + '\n'
    assert not any(secret in text for secret in secrets), 'an identity reached the fixture'
    OUT.write_text(text, encoding='utf-8')
    for case, answer, _ in record['cases']:
        print(case[:70], '->', answer)


if __name__ == '__main__':
    main()
