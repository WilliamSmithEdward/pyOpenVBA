"""What a call does to Err: what the called procedure reads of it, and what its caller reads afterwards, by how the call
is written and how the caller handles errors.

Each case raises error 1000 and makes the call. The called procedure records what it reads into gSeen, as its first
statement or after one of a few others. The answer is that, then what the caller reads straight after.

Rules this splits, with what each predicts for a Sub called by statement under Resume Next:
- calls preserve Err, as docs/vba_runtime.md said: seen 1000, after 1000.
- any call clears Err: seen 0, after 0.
- only a Function clears it: seen 1000, after 1000, where a Function case reads 0, 0.
- leaving the procedure clears it: seen 1000, after 0.

The first holds, until the called procedure runs an On Error statement. The harness's line numbers inject one into every
procedure, so with them on, as run_vba has them by default, every call here reads 0: this has to run with them off.

    python scripts/measure_err_on_call.py

writes tests/fixtures/err_on_call.json, which tests/test_vba_err_on_call.py replays.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/err_on_call.json'
CALLS = {
    'read directly': 'gSeen = "direct " & Err.Number',
    'sub': 'PlainSub',
    'call sub': 'Call PlainSub',
    'sub with brackets': 'Call PlainSub()',
    'public sub': 'PublicSub',
    'sub with an argument': 'ArgSub 1',
    'function alone': 'v = PlainFunction()',
    'function no brackets': 'v = PlainFunction',
    'function by statement': 'PlainFunction',
    'function in an expression': 'v = "x" & PlainFunction()',
    'function in an if': 'If PlainFunction() = "" Then v = 1',
    'public function': 'v = PublicFunction()',
    'function with an argument': 'v = ArgFunction(1)',
    'function given the number': 'v = ArgFunction(Err.Number)',
    'property': 'v = PlainProperty',
    'sub that traps': 'TrapSub',
    'sub that traps then clears': 'TrapClearSub',
    'sub that exits': 'ExitSub',
    'sub that errs and clears': 'FailSub',
    'run by name': 'Application.Run "PublicSub"',
    'builtin': 'v = Len("abc")',
    'host member': 'v = Application.Name',
    'dictionary': 'CreateObject("Scripting.Dictionary").Add "a", 1',
}
RAISES = {'raise': 'Err.Raise 1000, , "raised"', 'zero': 'v = 1 / gZero', 'object': 'gNothing.Name = 1'}
HANDLERS = ('resume next', 'handler')
CASES = [(handler, raises, name) for handler in HANDLERS for raises in RAISES for name in CALLS]

_HELPERS = '''
Private gSeen As String, gZero As Long, gNothing As Object

Private Sub Note(ByVal where As String)
    gSeen = where & " " & Err.Number
End Sub

Private Sub PlainSub()
    gSeen = "sub " & Err.Number
End Sub

Public Sub PublicSub()
    gSeen = "public " & Err.Number
End Sub

Private Sub ArgSub(ByVal a As Long)
    gSeen = "arg " & Err.Number
End Sub

Private Function PlainFunction() As String
    gSeen = "fn " & Err.Number
End Function

Public Function PublicFunction() As String
    gSeen = "publicfn " & Err.Number
End Function

Private Function ArgFunction(ByVal a As Long) As String
    gSeen = "argfn " & a & " " & Err.Number
End Function

Private Property Get PlainProperty() As String
    gSeen = "prop " & Err.Number
End Property

Private Sub TrapSub()
    On Error Resume Next
    gSeen = "trap " & Err.Number
End Sub

Private Sub TrapClearSub()
    On Error GoTo 0
    gSeen = "trap0 " & Err.Number
End Sub

Private Sub ExitSub()
    gSeen = "exit " & Err.Number
    Exit Sub
End Sub

Private Sub FailSub()
    On Error Resume Next
    gSeen = "fail " & Err.Number
    Err.Raise 2000
    gSeen = gSeen & " " & Err.Number
End Sub
'''


def _case(index: int, handler: str, raises: str, name: str) -> list[str]:
    fail, call = RAISES[raises], CALLS[name]
    lines = [f'Private Function Case{index}() As String', '    Dim v As Variant', '    gSeen = ""']
    if handler == 'resume next':
        lines += ['    On Error Resume Next', f'    {fail}', f'    {call}',
                  f'    Case{index} = gSeen & " after " & Err.Number']
    else:
        lines += ['    On Error GoTo Handler', f'    {fail}', f'    Case{index} = "no error"', '    Exit Function',
                  'Handler:', f'    {call}', f'    Case{index} = gSeen & " after " & Err.Number']
    return [*lines, 'End Function', '']


def _probe() -> str:
    lines = [_HELPERS]
    for index, (handler, raises, name) in enumerate(CASES):
        lines += _case(index, handler, raises, name)
    lines += ['Public Function Probe(ByVal first As Long, ByVal last As Long) As String',
              '    Dim i As Long, s As String',
              '    For i = first To last', '        Select Case i']
    lines += [f'            Case {index}: s = s & "@@@" & Case{index}()' for index in range(len(CASES))]
    lines += ['        End Select', '    Next', '    Probe = Mid$(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def main() -> None:
    source = _probe()
    answers: list[str] = []
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        for first in range(0, len(CASES), 40):
            last = min(first + 39, len(CASES) - 1)
            # Without the harness's line numbers: the handlers they inject run On Error, which clears Err on entry.
            result = excel.run_vba(source, 'Probe', (first, last), timeout=120.0, line_numbers=False)
            assert result.ok and not result.dialogs, result
            answers += str(result.value).split('@@@')
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source,
              'cases': [[handler, raises, name, answer]
                        for (handler, raises, name), answer in zip(CASES, answers, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for handler, raises, name, answer in record['cases']:
        print(f'{handler:12} {raises:7} {name:28} {answer}')


if __name__ == '__main__':
    main()
