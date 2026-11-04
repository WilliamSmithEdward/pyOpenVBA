"""Application.ExecuteExcel4Macro in live Excel, for what macros use it for: reading a cell of a workbook, open or
closed, in R1C1, and working out a plain expression. Every book is a temporary file, deleted afterwards. Answers name
the folder as <folder>, so no machine path reaches the fixture.

    python scripts/measure_execute_excel4.py

writes tests/fixtures/execute_excel4.json, which tests/test_excel_execute_excel4.py replays.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / "tests/fixtures/execute_excel4.json"

#: The text given, <folder> standing for the temporary folder and <here> for the open workbook's name.
ROWS = [
    "1+1", "2*3+4", '"a"&"b"', "TRUE", "FALSE", "1/0", "NA()", '"text"', "SUM(1,2,3)", "LEN(\"abc\")", "", " ", "1+",
    "NoSuchName", "NOSUCH(1)",
    "R1C1", "R2C1", "!R1C1", "RC", "R[1]C[1]", "A1", "Sheet1!R1C1", "Sheet1!R2C1", "Sheet1!A1", "'Sheet1'!R1C1",
    "Nowhere!R1C1", "SUM(Sheet1!R1C1:R2C1)", "Sheet1!R1C1:R2C1", "Sheet1!R9C9",
    "'[<here>]Sheet1'!R1C1", "'[<here>]Sheet1'!R2C1", "'[<here>]Sheet1'!R9C9",
    "'<folder>\\[closed.xlsx]Data'!R1C1", "'<folder>\\[closed.xlsx]Data'!R1C2", "'<folder>\\[closed.xlsx]Data'!R2C1",
    "'<folder>\\[closed.xlsx]Data'!R3C1", "'<folder>\\[closed.xlsx]Data'!R4C1", "'<folder>\\[closed.xlsx]Data'!R9C9",
    "'<folder>\\[closed.xlsx]Nowhere'!R1C1", "'<folder>\\[missing.xlsx]Data'!R1C1",
    "'<folder>\\[closed.xlsx]Data'!A1", "SUM('<folder>\\[closed.xlsx]Data'!R1C1:R1C1)",
]

PROBE = '''Function One(ByVal text As String) As String
    Dim v As Variant
    On Error GoTo Failed
    v = Application.ExecuteExcel4Macro(text)
    If IsError(v) Then
        One = "error:" & CStr(v)
    ElseIf IsObject(v) Then
        One = "object:" & TypeName(v)
    ElseIf IsArray(v) Then
        One = "array"
    Else
        One = TypeName(v) & ":" & CStr(v)
    End If
    Exit Function
Failed:
    One = "E" & Err.Number
End Function

Function Setup(ByVal folder As String) As String
    Application.DisplayAlerts = False
    Dim b As Workbook, here As Workbook
    Set here = ActiveWorkbook
    here.Worksheets(1).Range("A1").Value = 11
    here.Worksheets(1).Range("A2").Value = "two"
    Set b = Workbooks.Add
    b.Worksheets(1).Name = "Data"
    b.Worksheets(1).Range("A1").Value = 42
    b.Worksheets(1).Range("B1").Value = "word"
    b.Worksheets(1).Range("A2").Value = True
    b.Worksheets(1).Range("A3").Formula = "=A1*2"
    b.Worksheets(1).Range("A4").Value = DateSerial(2026, 1, 2)
    b.SaveAs folder & "\\closed.xlsx", 51
    b.Close False
    here.Activate
    Setup = here.Name
End Function
'''


def _call(text: str, folder: str, here: str) -> str:
    return text.replace("<folder>", folder).replace("<here>", here).replace('"', '""')


def main() -> None:
    import shutil
    import tempfile

    folder = tempfile.mkdtemp(prefix="pyopenvba_xl4_")
    rows: list[dict[str, str]] = []
    try:
        with ExcelSession(HarnessConfig(lock_wait_s=60.0)) as excel:
            excel.new_document()
            set_up = excel.run_vba(PROBE, "Setup", args=[folder], timeout=60.0)
            assert set_up.ok, set_up
            here = str(set_up.value)
            for text in ROWS:
                body = f'Function Run1() As String\nRun1 = One("{_call(text, folder, here)}")\nEnd Function\n'
                result = excel.run_vba(PROBE + body, "Run1", timeout=30.0)
                value = str(result.value) if result.ok else f"failed {result.outcome}"
                value = value.replace(folder, "<folder>")
                print(repr(text), value[:80], flush=True)
                rows.append({"text": text, "result": value})
    finally:
        shutil.rmtree(folder, ignore_errors=True)
    OUT.write_text(json.dumps({"measured_at": datetime.now(timezone.utc).isoformat(), "probe": PROBE, "rows": rows},
                              indent=1) + "\n", encoding="utf-8")


if __name__ == "__main__":
    main()
