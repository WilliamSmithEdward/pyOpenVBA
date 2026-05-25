"""Bake a single-module read/write diagnostic for Access GUI.

Confirms that modifying an EXISTING module's body (no rename, no
add, no delete) round-trips cleanly through Access. If this passes
where delete/add/rename crash, we know the in-place body rewrite
path is the only safe operation today.
"""

from __future__ import annotations

import shutil
from pathlib import Path

from pyopenvba.access import AccessFile

REPO = Path(__file__).resolve().parents[1]
CORPUS = REPO / "tests" / "live_access_test" / "re_corpus" / "samples"
OUT = REPO / "demo" / "output" / "access_phase5f"
SAMPLE_STD = CORPUS / "040__sub_msgbox_hello.accdb"


def _copy(name: str) -> Path:
    dst = OUT / name
    dst.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy(SAMPLE_STD, dst)
    return dst


def _build_100_line_body() -> str:
    lines: list[str] = [
        "Option Compare Database",
        "Option Explicit",
        "",
        "' 100-line read/write round-trip test for Access GUI.",
        "' Each Sub below should compile and run cleanly.",
        "",
        "Public Sub RunAll()",
        "    Greeting",
        "    LoopDemo",
        "    StringDemo",
        "    MathDemo",
        "    NestedIf",
        "    SelectCaseDemo",
        "    ArrayDemo",
        "    FinalReport",
        "End Sub",
        "",
        "Public Sub Greeting()",
        '    Debug.Print "Hello from pyOpenVBA Access write path."',
        "End Sub",
        "",
        "Public Sub LoopDemo()",
        "    Dim i As Long",
        "    Dim total As Long",
        "    total = 0",
        "    For i = 1 To 10",
        "        total = total + i",
        "    Next i",
        '    Debug.Print "LoopDemo sum 1..10 = " & total',
        "End Sub",
        "",
        "Public Sub StringDemo()",
        "    Dim s As String",
        '    s = "abcdef"',
        '    Debug.Print "StringDemo length = " & Len(s)',
        '    Debug.Print "StringDemo upper  = " & UCase(s)',
        "End Sub",
        "",
        "Public Sub MathDemo()",
        "    Dim x As Double",
        "    Dim y As Double",
        "    x = 3.14159",
        "    y = x * 2",
        '    Debug.Print "MathDemo 2*pi ~= " & y',
        "End Sub",
        "",
        "Public Sub NestedIf()",
        "    Dim n As Long",
        "    n = 7",
        "    If n > 0 Then",
        "        If n Mod 2 = 0 Then",
        '            Debug.Print "positive even"',
        "        Else",
        '            Debug.Print "positive odd"',
        "        End If",
        "    Else",
        '        Debug.Print "non-positive"',
        "    End If",
        "End Sub",
        "",
        "Public Sub SelectCaseDemo()",
        "    Dim k As Long",
        "    k = 2",
        "    Select Case k",
        "        Case 1",
        '            Debug.Print "one"',
        "        Case 2",
        '            Debug.Print "two"',
        "        Case 3",
        '            Debug.Print "three"',
        "        Case Else",
        '            Debug.Print "other"',
        "    End Select",
        "End Sub",
        "",
        "Public Sub ArrayDemo()",
        "    Dim arr(1 To 5) As Long",
        "    Dim i As Long",
        "    For i = 1 To 5",
        "        arr(i) = i * i",
        "    Next i",
        "    For i = 1 To 5",
        '        Debug.Print "arr(" & i & ") = " & arr(i)',
        "    Next i",
        "End Sub",
        "",
        "Public Sub FinalReport()",
        '    Debug.Print "------------------------------"',
        '    Debug.Print "FinalReport: all subs ran OK."',
        '    Debug.Print "------------------------------"',
        "End Sub",
        "",
        "' Sentinel comment to verify the bottom of the module survives.",
        "' If you can read this line in the Access VBA editor, the",
        "' 100-line write path round-tripped cleanly.",
        "",
    ]
    # Pad/truncate to exactly 100 lines of source body.
    while len(lines) < 100:
        lines.append("' filler line " + str(len(lines)))
    lines = lines[:100]
    return "\r\n".join(lines) + "\r\n"


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    dst = _copy("diag_K_100line_write.accdb")
    body = _build_100_line_body()
    line_count = body.count("\r\n")
    print(f"Writing {line_count} lines ({len(body)} chars) into module 'M' ...")

    db = AccessFile(dst)
    db.set_module("M", body)
    db.save()

    # Re-open and read back to confirm the Python round-trip.
    db2 = AccessFile(dst)
    readback = db2.get_module("M")
    rb_lines = readback.count("\r\n")
    print(f"Read back: {rb_lines} lines ({len(readback)} chars)")
    print(f"Body bytes round-trip exact match: {readback == body}")
    if readback != body:
        # Diff first divergence.
        for i, (a, b) in enumerate(zip(body, readback)):
            if a != b:
                print(f"  first diff at char {i}: {a!r} vs {b!r}")
                print(f"  context (orig):    {body[max(0, i - 20):i + 20]!r}")
                print(f"  context (readback):{readback[max(0, i - 20):i + 20]!r}")
                break
        else:
            print(f"  length differs: orig={len(body)} readback={len(readback)}")

    print()
    print("OPEN IN ACCESS VBE:")
    print(f"  {dst}")
    print("  - Modules > M should contain the 100-line body above.")
    print("  - Sentinel comment at the bottom should be present.")
    print("  - F5 inside RunAll() should print 8 sections to Immediate.")
    print("  - Debug > Compile should succeed without errors.")


if __name__ == "__main__":
    main()
