"""When VBA clears the Err object: on an On Error statement, on leaving a procedure, and what a call leaves behind.

Each case is a small program run in live Excel; its Probe function
reports Err.Number at the points that matter. The answers go to
tests/fixtures/vba_semantics/err_lifetime.json, which
tests/test_vba_err_lifetime.py holds the interpreter to.

    python scripts/measure_err_lifetime.py
"""

from __future__ import annotations

import json
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "tests" / "fixtures" / "vba_semantics" / "err_lifetime.json"

#: Procedures every case may call.
HELPERS = """Private Function RaisesHandledEnd() As String
    On Error Resume Next
    Err.Raise 1004
    RaisesHandledEnd = "inside " & Err.Number
End Function

Private Function RaisesHandledExit() As String
    On Error Resume Next
    Err.Raise 1004
    RaisesHandledExit = "inside " & Err.Number
    Exit Function
End Function

Private Sub RaisesHandledSub()
    On Error Resume Next
    Err.Raise 1004
End Sub

Private Function SetsWithoutHandler() As String
    Err.Number = 5
    SetsWithoutHandler = "inside " & Err.Number
End Function

Private Function Plain() As String
    Plain = "inside " & Err.Number
End Function

Private Function PlainExit() As String
    PlainExit = "inside " & Err.Number
    Exit Function
End Function

Private Function Clears() As String
    Err.Clear
    Clears = "inside " & Err.Number
End Function

Private Function RaisesUnhandled() As String
    Err.Raise 1004
End Function

Private Function HandlerLabel() As String
    On Error GoTo Bad
    Err.Raise 13
    HandlerLabel = "not here"
    Exit Function
Bad:
    HandlerLabel = "handled " & Err.Number
End Function

Private Function HandlerResume() As String
    Dim seen As Long
    On Error GoTo Bad
    Err.Raise 13
    HandlerResume = "resumed " & seen & " " & Err.Number
    Exit Function
Bad:
    seen = Err.Number
    Resume Next
End Function

Private Function ErrorStatementHandled() As String
    On Error Resume Next
    Error 11
    ErrorStatementHandled = "inside " & Err.Number
End Function

Private Function HandlerExit() As String
    On Error GoTo Bad
    Err.Raise 13
    Exit Function
Bad:
    HandlerExit = "handled " & Err.Number
    Exit Function
End Function

Private Sub PlainSub()
    Dim x As Long
    x = 1
End Sub

Private Function Outer() As String
    Outer = "outer sees " & Err.Number & ", " & Plain()
End Function
"""

#: Each case: its name, and the body of a function that answers with what it saw.
CASES: list[tuple[str, str]] = [
    ("handled_end", 'out = RaisesHandledEnd() & " / after " & Err.Number'),
    ("handled_exit", 'out = RaisesHandledExit() & " / after " & Err.Number'),
    ("handled_sub", 'RaisesHandledSub\nout = "after " & Err.Number'),
    ("set_without_handler", 'out = SetsWithoutHandler() & " / after " & Err.Number'),
    ("caller_error_through_plain_call",
     'On Error Resume Next\nErr.Raise 7\nout = "before " & Err.Number & " / " & Plain() & " / after " & Err.Number'),
    ("caller_error_through_plain_exit",
     'On Error Resume Next\nErr.Raise 7\nout = "before " & Err.Number & " / " & PlainExit() & " / after " & Err.Number'),
    ("caller_error_through_handled_call",
     'On Error Resume Next\nErr.Raise 7\nout = "before " & Err.Number & " / " & RaisesHandledEnd() & " / after " & '
     "Err.Number"),
    ("callee_clears", 'On Error Resume Next\nErr.Raise 7\nout = Clears() & " / after " & Err.Number'),
    ("unhandled_reaches_caller", 'On Error Resume Next\nout = "skipped"\nout = RaisesUnhandled()\n'
                                 'out = out & " / after " & Err.Number'),
    ("on_error_again", 'On Error Resume Next\nErr.Raise 7\nOn Error Resume Next\nout = "after " & Err.Number'),
    ("on_error_goto_0", 'On Error Resume Next\nErr.Raise 7\nOn Error GoTo 0\nout = "after " & Err.Number'),
    ("on_error_goto_label", 'On Error Resume Next\nErr.Raise 7\nOn Error GoTo Bad\nout = "after " & Err.Number\n'
                            'GoTo Done\nBad:\nout = "bad"\nDone:'),
    ("handler_exit", 'out = HandlerExit() & " / after " & Err.Number'),
    ("caller_error_through_builtin", 'On Error Resume Next\nErr.Raise 7\nDim n As Long\nn = Len("abc")\n'
                                     'out = "after " & Err.Number'),
    ("caller_error_through_object", 'On Error Resume Next\nErr.Raise 7\nDim s As String\ns = Application.Name\n'
                                    'out = "after " & Err.Number'),
    ("caller_error_through_sub", 'On Error Resume Next\nErr.Raise 7\nPlainSub\nout = "after " & Err.Number'),
    ("set_then_call", 'Err.Number = 5\nout = "before " & Err.Number & " / " & Plain() & " / after " & Err.Number'),
    ("nested_calls", 'On Error Resume Next\nErr.Raise 7\nout = Outer() & " / after " & Err.Number'),
    ("handler_label", 'out = HandlerLabel() & " / after " & Err.Number'),
    ("handler_resume", 'out = HandlerResume() & " / after " & Err.Number'),
    ("error_statement", 'out = ErrorStatementHandled() & " / after " & Err.Number'),
    ("resume_next_keeps", 'On Error Resume Next\nErr.Raise 7\nDim x As Long\nx = 1\nout = "after " & Err.Number'),
    ("second_error_replaces", 'On Error Resume Next\nErr.Raise 7\nErr.Raise 9\nout = "after " & Err.Number'),
]


def module() -> str:
    lines = [HELPERS]
    for index, (_, body) in enumerate(CASES):
        lines += [f"Private Function C{index}() As String", "Dim out As String", body, f"C{index} = out",
                  "End Function"]
    lines += ["Public Function Probe() As String", "Dim out As String"]
    lines += [f'out = out & "{name}=" & C{index}() & "|"' for index, (name, _) in enumerate(CASES)]
    lines += ["Probe = out", "End Function"]
    return "\n".join(lines) + "\n"


def main() -> None:
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(module(), "Probe", timeout=300.0)
        assert result.ok, f"{result.outcome}: {result.message} {result.error}"
    answers = dict(part.split("=", 1) for part in str(result.value).split("|") if "=" in part)
    record = {"helpers": HELPERS, "cases": [{"name": name, "body": body, "answer": answers[name]}
                                            for name, body in CASES]}
    OUT.write_text(json.dumps(record, indent=1) + "\n", encoding="utf-8")
    for case in record["cases"]:
        print(f"{case['name']:36} {case['answer']}")


if __name__ == "__main__":
    main()
