"""What a Set into a variable of a class answers when the object is of another class, each line compiled alone in live Excel.

The variable typed and the object typed, the object bound late into a class, a class into Object and Variant, default members, Nothing, New,
For Each, a typed parameter and a function's typed answer. No file, dialog or other program is touched.

    python scripts/measure_typed_set.py

writes tests/fixtures/typed_set.json, keeping each line as it finishes and asking only the rest.
"""
from __future__ import annotations

import json
from pathlib import Path

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/typed_set.json'
#: Each line, and what it reads back: the variable it set.
LINES: list[tuple[str, str]] = [
    ('Set c = Worksheets(1)', 'c'), ('Set c = ws', 'c'), ('Set c = r', 'c'), ('Set c = wb', 'c'), ('Set c = o', 'c'), ('Set c = d', 'c'),
    ('Set c = Worksheets', 'c'), ('Set c = Nothing', 'c'), ('Set c = New Collection', 'c'), ('Set c = Application', 'c'),
    ('Set ws = Worksheets(1)', 'ws'), ('Set ws = ActiveSheet', 'ws'), ('Set ws = Sheets(1)', 'ws'), ('Set ws = Range("A1")', 'ws'), ('Set ws = r', 'ws'),
    ('Set ws = wb', 'ws'), ('Set ws = o', 'ws'), ('Set ws = oc', 'ws'), ('Set ws = Nothing', 'ws'), ('Set ws = ThisWorkbook.Sheets(1)', 'ws'),
    ('Set ws = ws.Parent', 'ws'), ('Set ws = r.Worksheet', 'ws'), ('Set ws = r.Parent', 'ws'),
    ('Set r = ws', 'r'), ('Set r = ws.Range("A1:B2").Cells', 'r'), ('Set r = ws.Rows(1)', 'r'), ('Set r = ws.UsedRange', 'r'), ('Set r = o', 'r'),
    ('Set r = ws.Range("A1").Value', 'r'), ('Set r = ws.Cells(1, 1)', 'r'), ('Set r = Selection', 'r'), ('Set r = ws.Shapes', 'r'),
    ('Set wb = ws', 'wb'), ('Set wb = ws.Parent', 'wb'), ('Set wb = Workbooks(1)', 'wb'), ('Set wb = o', 'wb'),
    ('Set w = ws', 'w'), ('Set w = r', 'w'), ('Set w = c', 'w'),
    ('Set o = ws', 'o'), ('Set o = c', 'o'), ('Set o = Nothing', 'o'), ('Set v = c', 'v'), ('Set v = ws', 'v'), ('Set v = 5', 'v'), ('Set o = 5', 'o'),
    ('Set c = oc', 'c'), ('Set c = vc', 'c'), ('Set c = vw', 'c'), ('Set ws = vw', 'ws'),
    ('Takes o', 'c'), ('Takes oc', 'c'), ('TakesRef o', 'c'), ('Set c = Gives()', 'c'), ('Set c = GivesSheet()', 'c'), ('Set ws = GivesSheet()', 'ws'),
    ('For Each c In Worksheets: Next', 'c'), ('For Each ws In Worksheets: Next', 'ws'), ('For Each r In ws.Range("A1:A2"): Next', 'r'),
]
SOURCE = '''Private Sub Takes(ByVal x As Collection)
End Sub
Private Sub TakesRef(ByRef x As Collection)
End Sub
Private Function Gives() As Collection
    Set Gives = Worksheets(1)
End Function
Private Function GivesSheet() As Object
    Set GivesSheet = Worksheets(1)
End Function
Function Probe() As String
    Dim c As Collection, ws As Worksheet, r As Range, wb As Workbook, w As Excel.Worksheet, o As Object, oc As Object, v As Variant, vc As Variant, vw As Variant, d As Object
    Set ws = Worksheets(1): Set r = ws.Range("A1"): Set wb = ThisWorkbook: Set o = ws: Set oc = New Collection: Set vc = New Collection: Set vw = ws
    Set d = CreateObject("Scripting.Dictionary")
    Dim e As Long
    On Error Resume Next
    LINE
    e = Err.Number
    Err.Clear
    Probe = e & "=" & TypeName(READ)
End Function
'''


def macro(line: str, read: str) -> str:
    return SOURCE.replace('LINE', line.replace(': ', '\n    ')).replace('READ', read)


def main() -> None:
    from pyvbaharness import ExcelSession, HarnessConfig
    from pyvbaharness.results import HarnessError

    answers: dict[str, str] = json.loads(OUT.read_text(encoding='utf-8'))['answers'] if OUT.exists() else {}
    left = [pair for pair in LINES if pair[0] not in answers]
    tries = 0
    while left and tries < 12:
        tries += 1
        # The worker can drop out between compile errors. What it had answered is kept and the rest asked of a new one.
        try:
            with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
                while left:
                    line, read = left[0]
                    excel.new_document()
                    result = excel.run_vba(macro(line, read), 'Probe', timeout=40.0, line_numbers=False)
                    if result.dialogs:
                        answers[line] = 'does not compile: ' + ' '.join(result.dialogs[0].message.split())
                    elif result.ok:
                        answers[line] = str(result.value)
                    else:
                        raise SystemExit(f'{line} stopped: {result.message}')
                    left.pop(0)
                    OUT.write_text(json.dumps({'source': SOURCE, 'answers': answers}, indent=1) + '\n', encoding='utf-8', newline='\n')
        except HarnessError as failed:
            print('worker dropped at', left[0][0], str(failed)[:60], flush=True)
    print(len(answers), 'of', len(LINES))


if __name__ == '__main__':
    main()
