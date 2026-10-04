"""Record isolated GoSub/Return probes in real Excel for offline conformance."""

from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

import pyvbaharness as harness

DESTINATION = Path(__file__).resolve().parents[1] / "tests/fixtures/vba_semantics/gosub.json"


def sources() -> dict[str, str]:
    bodies = {
        "nested": '''Dim s As String
s = "A"
GoSub First
Probe = s & "Z"
Exit Function
First:
s = s & "B"
GoSub Second
s = s & "D"
Return
Second:
s = s & "C"
Return''',
        "loop": '''Dim i As Long, s As String
For i = 1 To 3
    GoSub Append
    s = s & "!"
Next
Probe = s
Exit Function
Append:
s = s & CStr(i)
Return''',
        "goto_in_subroutine": '''GoSub First
Probe = Probe & "Z"
Exit Function
First:
Probe = "A"
GoTo Second
Second:
Probe = Probe & "B"
Return''',
        "early_exit": '''GoSub First
Probe = "wrong"
Exit Function
First:
Probe = "exit"
Exit Function''',
        "return_without_gosub": '''On Error Resume Next
Return
Probe = CStr(Err.Number)''',
        "resume_next": '''On Error Resume Next
GoSub First
Probe = Probe & "Z"
Exit Function
First:
Err.Raise 11
Probe = CStr(Err.Number)
Return''',
        "handler_resume": '''On Error GoTo Handler
GoSub First
Probe = Probe & "Z"
Exit Function
First:
Err.Raise 11
Probe = Probe & "B"
Return
Handler:
Probe = CStr(Err.Number)
Resume Next''',
        "handler_return": '''On Error GoTo Handler
GoSub First
Probe = Probe & "Z:" & CStr(Err.Number)
Exit Function
First:
Err.Raise 11
Probe = "wrong"
Return
Handler:
Probe = CStr(Err.Number)
Return''',
        "numeric_label": '''GoSub 100
Probe = Probe & "Z"
Exit Function
100 Probe = "N"
Return''',
        "fall_off_procedure": '''GoSub Last
Probe = "wrong"
Exit Function
Last:
Probe = "end"''',
        "conditional_return": '''Dim n As Long
GoSub First
GoSub First
Probe = CStr(n)
Exit Function
First:
n = n + 1
If n = 1 Then
    Return
End If
n = n + 10
Return''',
        "gosub_from_handler": '''On Error GoTo Handler
Err.Raise 11
Probe = Probe & "Z"
Exit Function
Worker:
Probe = Probe & "W"
Return
Handler:
Probe = "H"
GoSub Worker
Resume Next''',
        "returned_stack_empty": '''On Error Resume Next
GoSub Worker
Return
Probe = CStr(Err.Number)
Exit Function
Worker:
Return''',
        "goto_after_error": '''On Error GoTo Handler
GoSub Worker
Probe = Probe & "Z"
Exit Function
Worker:
Err.Raise 11
Return
Handler:
Probe = "H"
GoTo Finish
Finish:
Probe = Probe & "F"
Return''',
    }
    for mode in ("GoSub", "GoTo"):
        for expression in ("-1", "0", "1", "2", "3", "255", "256", "0.5", "1.5", "2.5", "Null", '"2"', "32768", "True"):
            tail = "Return" if mode == "GoSub" else 'Probe = Probe & "Z": Exit Function'
            bodies[f"on_{mode.lower()}_{expression}"] = f'''On Error Resume Next
On {expression} {mode} First, Second
Probe = Probe & "Z:" & CStr(Err.Number)
Exit Function
First:
Probe = "A"
{tail}
Second:
Probe = "B"
{tail}'''
    return {name: "Function Probe() As String\n" + body + "\nEnd Function\n" for name, body in bodies.items()}


def main() -> None:
    rows = []
    with harness.ExcelSession(harness.HarnessConfig(lock_wait_s=30.0)) as office:
        office.new_document()
        metadata = office.run_vba('Function Metadata() As String\nMetadata = Application.Version & "|" & Application.Build & "|" & Application.International(1)\nEnd Function', "Metadata", timeout=30.0)
        if not metadata.ok:
            raise RuntimeError(metadata)
        for name, source in sources().items():
            result = office.run_vba(source, "Probe", timeout=30.0)
            if not result.ok:
                raise RuntimeError(f"{name}: {result}")
            rows.append({"name": name, "source": source, "result": result.value})
            print(f"{name}: {result.value}", flush=True)
    DESTINATION.write_text(json.dumps({"measured_at": datetime.now(timezone.utc).isoformat(), "excel_version_build_country": metadata.value, "probes": rows}, indent=2) + "\n", encoding="utf-8")


if __name__ == "__main__":
    main()
