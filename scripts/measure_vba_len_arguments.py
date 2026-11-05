"""Which arguments Len and LenB compile with and what they answer, each line compiled alone in live Excel.

    python scripts/measure_vba_len_arguments.py

writes tests/fixtures/vba_len_arguments.json, keeping each line as it finishes and asking only the rest.
"""
from __future__ import annotations

import json
from pathlib import Path

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/vba_len_arguments.json'
LINES = [
    'Len(5)', 'LenB(5)', 'Len(5.5)', 'Len(5&)', 'Len(5#)', 'Len(True)', 'Len(#1/2/2020#)', 'Len(&H10)', 'Len(-5)', 'Len(1 + 1)', 'Len((5))',
    'Len(K)', 'Len(KS)', 'Len(vbTab)', 'Len(xlUp)', 'Len(vbLong)', 'Len(n)', 'LenB(n)', 'Len(n + 1)', 'Len((n))', 'Len(-n)', 'Len(CLng(n))', 'Len(CStr(n))',
    'Len(d)', 'Len(b)', 'Len(dt)', 'Len(c)', 'Len(sg)', 'Len(by)', 'Len(i)', 'Len(ll)', 'Len(st)', 'Len(v)', 'Len(v + 1)', 'Len(n & "")', 'Len(Fn())', 'Len(FnS())', 'Len(FnV())',
    'Len(Null)', 'Len(Empty)', 'Len(ar(0))', 'LenB(ar(0))', 'Len(ty)', 'Len(ty.a)', 'Len(ty.b)', 'Len(md)', 'Len(ob)', 'Len(ws)', 'Len(ws.Name)', 'Len(ws.Index)', 'Len(ws.Range("A1"))',
    'Len(ws.Range("A1").Value)', 'Len(ws.Range("A1").Row)', 'Len(ob2.Index)', 'Len(Nothing)', 'Len(5 = 5)', 'Len(n = 5)', 'Len(v = 5)', 'Len(Date)', 'Len(Now)', 'Len(Timer)',
    'Len(Rows.Count)', 'LenB(Rows.Count)', 'Len(Me)',
]
SOURCE = '''Private Type T
    a As Long
    b As Integer
End Type
Const K As Long = 5
Const KS As String = "abc"
Private md As Double
Private Function Fn() As Long
    Fn = 12
End Function
Private Function FnS() As String
    FnS = "abc"
End Function
Private Function FnV() As Variant
    FnV = 123
End Function
Function Probe() As String
    Dim n As Long, d As Double, b As Boolean, dt As Date, c As Currency, sg As Single, by As Byte, i As Integer, ll As LongLong, st As String, v As Variant
    Dim ar(1) As Long, ty As T, ob As Object, ob2 As Object, ws As Worksheet, x As Variant
    n = 12345: d = 1.5: st = "abcd": v = 123: ar(0) = 7
    Set ws = ActiveSheet
    Set ob2 = ws
    ws.Range("A1").Value = 12345
    On Error Resume Next
    x = LINE
    Probe = Err.Number & "=" & TypeName(x) & ":" & x
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
