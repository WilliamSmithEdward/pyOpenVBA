"""Measure error handling: an error passing up through callers and class members,
Resume to a label and Resume Next, an error inside a handler, handlers in a loop,
Err.Raise with and without its arguments and out of range, Err's properties, the
Error function and Err.LastDllError.

Each probe is a module whose ``Function Probe() As String`` the fixture records
(see vba_probe_fixture.py); the class module is added to every session. A bare
Resume that retries an Err.Raise inside a one-line If loops forever natively as
well, so it is not measured.
"""
from __future__ import annotations
from pathlib import Path
from vba_probe_fixture import record

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/vba_error_handling.json'

CLASSES = {
    'CThrower': '''Option Explicit
Public Sub Boom()
    Err.Raise vbObjectError + 7, "CThrower.Boom", "kaboom"
End Sub
Public Function Divide(ByVal a As Double, ByVal b As Double) As Double
    Divide = a / b
End Function
Public Function Guarded() As String
    On Error GoTo Bad
    Err.Raise 5
    Exit Function
Bad:
    Guarded = "guarded " & Err.Number
End Function
''',
}

HELPERS = '''
Private Function Inner() As Long
    Inner = 1 / 0
End Function

Private Function Middle() As String
    Dim x As Long
    x = Inner()
    Middle = "not here"
End Function

Private Function ToLabel() As String
    On Error GoTo Handler
    Err.Raise 11
    ToLabel = "skipped"
Continue:
    ToLabel = ToLabel & "continued " & Err.Number
    Exit Function
Handler:
    ToLabel = "handled;"
    Resume Continue
End Function

Private Function ErrorInHandler() As String
    On Error GoTo H
    Err.Raise 5
    Exit Function
H:
    ErrorInHandler = "in handler;"
    Err.Raise 6
End Function

Private Function LoopWithHandler() As String
    Dim i As Long, s As String
    For i = 1 To 4
        On Error Resume Next
        Dim x As Double
        x = 10 / (i - 2)
        If Err.Number <> 0 Then s = s & "e" & Err.Number & ";" Else s = s & x & ";"
        Err.Clear
        On Error GoTo 0
    Next
    LoopWithHandler = s
End Function

Private Function NestedHandlers() As String
    On Error GoTo Outer
    Dim s As String
    s = Middle()
    NestedHandlers = s
    Exit Function
Outer:
    NestedHandlers = "outer caught " & Err.Number & " from " & Err.Source
End Function
'''

CASES = {
    'propagate_to_caller': 'On Error Resume Next\nDim s As String\ns = Middle()\nProbe = Err.Number & "|" & s & "|" & Err.Description',
    'resume_label': 'Probe = ToLabel()',
    'error_in_handler': 'On Error Resume Next\nDim s As String\ns = ErrorInHandler()\nProbe = Err.Number & "|" & s',
    'loop_with_handler': 'Probe = LoopWithHandler()',
    'nested_handlers': 'Probe = NestedHandlers()',
    'class_raise_custom': 'Dim c As New CThrower\nOn Error Resume Next\nc.Boom\nProbe = (Err.Number - vbObjectError) & "|" & Err.Source & "|" & Err.Description',
    'class_runtime_error_source': 'Dim c As New CThrower\nOn Error Resume Next\nDim d As Double\nd = c.Divide(1, 0)\nProbe = Err.Number & "|" & Err.Source',
    'class_guarded': 'Dim c As New CThrower\nProbe = c.Guarded() & "|" & Err.Number',
    'err_raise_args': 'On Error Resume Next\nErr.Raise 5, "MySrc", "My text", "help.chm", 42\nProbe = Err.Number & "|" & Err.Source & "|" & Err.Description & "|" & Err.HelpFile & "|" & Err.HelpContext',
    'err_raise_defaults': 'On Error Resume Next\nErr.Raise 9\nProbe = Err.Description & "|" & Err.Source & "|" & Err.HelpContext',
    'err_raise_zero': 'On Error Resume Next\nErr.Raise 0\nProbe = Err.Number & "|" & Err.Description',
    'err_raise_big': 'On Error Resume Next\nErr.Raise 70000\nProbe = Err.Number & "|" & Err.Description',
    'err_number_set': 'On Error Resume Next\nErr.Number = 13\nProbe = Err.Number & "|" & Err.Description\nErr.Description = "custom"\nProbe = Probe & "|" & Err.Description & "|" & Err.Number',
    'err_clear_props': 'On Error Resume Next\nErr.Raise 5, "S", "D"\nErr.Clear\nProbe = Err.Number & "|" & Err.Source & "|" & Err.Description & "|" & Len(Err.Source)',
    'resume_next_after_handler': 'On Error GoTo H\nDim s As String\nErr.Raise 5\ns = s & "after;"\nProbe = s & Err.Number\nExit Function\nH:\ns = "h;"\nResume Next',
    'on_error_goto_0_after_resume_next': 'On Error Resume Next\nErr.Raise 5\nOn Error GoTo 0\nProbe = Err.Number',
    'error_function_no_arg': 'On Error Resume Next\nErr.Raise 11\nProbe = Error() & "|" & Error$ & "|" & Error(Err.Number)',
    'cverr_vs_err': 'Dim v\nv = CVErr(5)\nProbe = (v = CVErr(5)) & "|" & IsError(v) & "|" & Err.Number',
    'err_lastdllerror': 'Probe = Err.LastDllError & "|" & TypeName(Err.LastDllError)',
}


def macro(name: str) -> str:
    return HELPERS + '\nFunction Probe() As String\n' + CASES[name] + '\nEnd Function\n'


def main() -> None:
    record({name: macro(name) for name in CASES}, OUT, CLASSES)


if __name__ == '__main__':
    main()
