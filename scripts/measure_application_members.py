"""Everyday Application members in live Excel: PathSeparator, Value, OperatingSystem, MeasurementUnit,
InchesToPoints, CentimetersToPoints, SheetsInNewWorkbook and ThisCell.

Each conversion's answer is compared, in VBA's own Double arithmetic, with the formulas that could give it, so the
fixture says which one Excel uses without spelling seventeen digits. SheetsInNewWorkbook is set, read back and
used by Workbooks.Add, and always put back to what it was: it is a setting Excel keeps for its user. ThisCell and
Caller are read from VBA Functions that formulas call, one cell's and an array formula's, and from a macro.

    python scripts/measure_application_members.py

writes tests/fixtures/application_members.json, which tests/test_excel_application_members.py replays.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / "tests/fixtures/application_members.json"

#: The numbers whose conversions are measured to the bit.
VALUES = [1, 3, -3, 1.5, -1.5, 7, -7, 0.1, -0.1, 1 / 3, -1 / 3, 5, -5, 12.7, -12.7, 100, -100, 2.2, -2.2, 0.7, -0.7,
          2.54, 10, 0.25, 6, 9, 11, 13, 0.01, 123.456, 1e300, -1e-300]

PROBE = r'''Private Type DoubleBits
    v As Double
End Type

Private Type ByteBits
    b(7) As Byte
End Type

Private Function Bits(r As Double) As String
    Dim d As DoubleBits, e As ByteBits, i As Long, s As String
    d.v = r
    LSet e = d
    For i = 7 To 0 Step -1
        s = s & Right$("0" & Hex$(e.b(i)), 2)
    Next
    Bits = s
End Function

Public Function Conversions(values As String) As String
    Dim s As String, x As Variant
    For Each x In Split(values, ";")
        s = s & "|" & Bits(Application.CentimetersToPoints(CDbl(x))) & "," & Bits(Application.InchesToPoints(CDbl(x)))
    Next
    Conversions = Mid$(s, 2)
End Function

Public Function WhoAmI() As String
    Dim caller As Variant
    caller = TypeName(Application.Caller)
    If caller = "Range" Then caller = Application.Caller.Address(False, False)
    WhoAmI = Application.ThisCell.Address(False, False) & "/" & caller
End Function

Private Function Tried(what As String, x As Variant) As String
    Dim r As Double
    On Error Resume Next
    If what = "in" Then r = Application.InchesToPoints(x) Else r = Application.CentimetersToPoints(x)
    If Err.Number <> 0 Then
        Tried = "E" & Err.Number
        Exit Function
    End If
    On Error GoTo 0
    Tried = CStr(r) & " " & TypeName(r)
End Function

Private Function Shown(v As Variant) As String
    Shown = CStr(v)
End Function

Private Function SetSheets(value As Variant) As String
    On Error Resume Next
    Application.SheetsInNewWorkbook = value
    SetSheets = "E" & Err.Number & ":" & Err.Description & ":" & Application.SheetsInNewWorkbook
    Err.Clear
End Function

Public Function Probe() As String
    Dim s As String, x As Variant, before As Long, wb As Workbook, ws As Worksheet, names As String
    s = "PathSeparator=" & Application.PathSeparator & " " & TypeName(Application.PathSeparator)
    s = s & "|Value=" & Application.Value & " " & TypeName(Application.Value)
    s = s & "|OperatingSystem=" & Application.OperatingSystem
    s = s & "|MeasurementUnit=" & Application.MeasurementUnit & " " & TypeName(Application.MeasurementUnit)
    For Each x In Array(1, 0.5, 2.54, -3, 0, 1 / 3, 1E+300, "2", True, "x")
        s = s & "|in " & CStr(x) & "=" & Tried("in", x)
    Next
    For Each x In Array(1, 2.54, 0.1, -3, 0, 10, 1 / 3, 1E+300, "2", True, "x")
        s = s & "|cm " & CStr(x) & "=" & Tried("cm", x)
    Next
    s = s & "|in empty=" & Tried("in", Empty) & "|in null=" & Tried("in", Null)
    before = Application.SheetsInNewWorkbook
    On Error GoTo Restore
    s = s & "|sheets type=" & TypeName(Application.SheetsInNewWorkbook)
    s = s & "|set 3=" & SetSheets(3)
    Set wb = Workbooks.Add
    names = ""
    For Each ws In wb.Worksheets
        names = names & "," & ws.Name
    Next
    s = s & "|add=" & wb.Worksheets.Count & Mid$(names, 1)
    wb.Close False
    Set wb = Workbooks.Add(xlWBATWorksheet)
    s = s & "|add worksheet=" & wb.Worksheets.Count
    wb.Close False
    s = s & "|set 0=" & SetSheets(0) & "|set 256=" & SetSheets(256) & "|set 255=" & SetSheets(255)
    s = s & "|set 1.6=" & SetSheets(1.6) & "|set 2.5=" & SetSheets(2.5) & "|set ""4""=" & SetSheets("4")
    s = s & "|set x=" & SetSheets("x") & "|set -1=" & SetSheets(-1) & "|set True=" & SetSheets(True)
Restore:
    Application.SheetsInNewWorkbook = before
    s = s & "|restored=" & (Application.SheetsInNewWorkbook = before)
    If Err.Number <> 0 Then s = s & "|error " & Err.Number
    On Error GoTo 0
    ' The formulas go in the probe's own workbook, whose module holds WhoAmI.
    Set ws = ThisWorkbook.Worksheets(1)
    ws.Range("A1").Formula = "=WhoAmI()"
    ws.Range("B1:B3").FormulaArray = "=WhoAmI()"
    ws.Range("C5").Formula = "=LEN(WhoAmI())"
    s = s & "|cell=" & Shown(ws.Range("A1").Value) & "|array=" & Shown(ws.Range("B1").Value) & "," & _
        Shown(ws.Range("B2").Value) & "," & Shown(ws.Range("B3").Value) & "|inside=" & Shown(ws.Range("C5").Value)
    On Error Resume Next
    x = Application.ThisCell.Address
    s = s & "|macro=E" & Err.Number & ":" & Err.Description
    Err.Clear
    On Error GoTo 0
    ws.Range("A1:C5").Clear
    Probe = s
End Function
'''


def main() -> None:
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        version = excel.run_vba('Function V() As String\nV = Application.Version & "|" & Application.Build\n'
                                'End Function', "V")
        result = excel.run_vba(PROBE, "Probe", timeout=300.0)
        assert result.ok and not result.dialogs, f"{result.outcome}: {result.message} {result.error}"
        bits = excel.run_vba(PROBE, "Conversions", [";".join(repr(value) for value in VALUES)], timeout=300.0)
        assert bits.ok and not bits.dialogs, f"{bits.outcome}: {bits.message} {bits.error}"
    answers = str(result.value).split("|")
    # Each value's CentimetersToPoints and InchesToPoints, to the bit, as big-endian hex.
    conversions = dict(zip((repr(value) for value in VALUES), str(bits.value).split("|"), strict=True))
    OUT.write_text(json.dumps({"excel_version_build": version.value,
                               "measured_at": datetime.now(timezone.utc).isoformat(),
                               "probe": PROBE, "answers": answers, "conversions": conversions}, indent=2) + "\n",
                   encoding="utf-8")
    for answer in answers:
        print(answer)
    print(conversions)


if __name__ == "__main__":
    main()
