"""Which calls compile with an empty argument at the end, f(a, ), or in the middle, f(a, , c), each line compiled alone in live Excel.

Built-in functions, a worksheet function and Range's members typed and late, the macro's own procedures as a function, with Call and as a statement.
No dialog is asked for and no file named.

    python scripts/measure_trailing_arguments.py

writes tests/fixtures/trailing_arguments.json, keeping each line as it finishes and asking only the rest.
"""
from __future__ import annotations

import json
from pathlib import Path

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/trailing_arguments.json'
LINES = [
    'x = Fun(1, )', 'x = Fun(1,)', 'x = Fun(1, , )', 'x = Fun(, )', 'x = Fun(1, , 3)', 'x = Fun(, 2)', 'x = Fun(1)',
    'Call Pro(1, )', 'Call Pro(1, , 3)', 'Pro 1, ', 'Pro 1, , 3', 'Pro 1,', 'x = Two(1, )',
    'x = Mid("abc", 1, )', 'x = Mid("abc", , 2)', 'x = InStr(1, "abc", )', 'x = Left("abc", )', 'x = Format(1, )', 'x = Array(1, )', 'x = Array(1, , 3)',
    'x = Replace("abc", "b", "x", )', 'x = IIf(True, 1, )',
    'x = WorksheetFunction.XLookup(1, r, r, )', 'x = WorksheetFunction.XLookup(1, r, r, , 0)', 'x = WorksheetFunction.Sum(1, )', 'x = WorksheetFunction.Sum(1, , 2)',
    'x = w.XLookup(1, r, r, )', 'x = w.XLookup(1, r, r, , 0)', 'x = w.Sum(1, )',
    'x = ws.Range("A1", ).Address', 'x = ws.Cells(1, ).Address', 'Set x = r.Find("x", )', 'Set x = r.Find("x", , xlValues)',
    'x = wo.Range("A1", ).Address', 'x = wo.Cells(1, ).Address', 'Set x = ro.Find("x", )', 'Call r.Find("x", )', 'r.ClearContents ', 'r.Sort r, ',
    'x = CallByName(r, "Address", VbGet, )',
]
SOURCE = '''Private Heard As String
Private Function Fun(ByVal a As Variant, Optional ByVal b As Variant, Optional ByVal c As Variant) As String
    Fun = IsMissing(a) & IsMissing(b) & IsMissing(c)
End Function
Private Function Two(ByVal a As Variant, ByVal b As Variant) As String
    Two = "two"
End Function
Private Sub Pro(ByVal a As Variant, Optional ByVal b As Variant, Optional ByVal c As Variant)
    Heard = IsMissing(a) & IsMissing(b) & IsMissing(c)
End Sub
Function Probe() As String
    Dim x As Variant, ws As Worksheet, r As Range, w As WorksheetFunction, wo As Object, ro As Object
    Set ws = ActiveSheet: Set r = ws.Range("A1:A3"): Set w = WorksheetFunction: Set wo = ws: Set ro = r
    r.Value = 1
    On Error Resume Next
    LINE
    Dim e As Long, shown As String
    e = Err.Number
    Err.Clear
    If IsObject(x) Then
        shown = ""
    ElseIf IsArray(x) Then
        shown = LBound(x) & " to " & UBound(x) & ", " & TypeName(x(1))
    Else
        shown = CStr(x)
    End If
    Probe = e & "=" & TypeName(x) & ":" & shown & ":" & Heard
End Function
'''


def macro(line: str) -> str:
    return SOURCE.replace('LINE', line)


def main() -> None:
    from pyvbaharness import ExcelSession, HarnessConfig
    from pyvbaharness.results import HarnessError

    answers: dict[str, str] = json.loads(OUT.read_text(encoding='utf-8'))['answers'] if OUT.exists() else {}
    left = [line for line in LINES if line not in answers]
    tries = 0
    while left and tries < 12:
        tries += 1
        # The worker can drop out between compile errors. What it had answered is kept and the rest asked of a new one.
        try:
            with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
                while left:
                    line = left[0]
                    excel.new_document()
                    result = excel.run_vba(macro(line), 'Probe', timeout=40.0, line_numbers=False)
                    if result.dialogs:
                        answers[line] = 'does not compile: ' + ' '.join(result.dialogs[0].message.split())
                    elif result.ok:
                        answers[line] = str(result.value)
                    else:
                        raise SystemExit(f'{line} stopped: {result.message}')
                    left.pop(0)
                    OUT.write_text(json.dumps({'source': SOURCE, 'answers': answers}, indent=1) + '\n', encoding='utf-8', newline='\n')
        except HarnessError as failed:
            print('worker dropped at', left[0], str(failed)[:60], flush=True)
    print(len(answers), 'of', len(LINES))


if __name__ == '__main__':
    main()
