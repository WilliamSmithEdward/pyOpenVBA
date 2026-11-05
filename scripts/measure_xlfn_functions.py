"""Each of Excel's functions once, by the calls tests/fixtures/formula_prefixes/prefixes.json has: which hidden names a formula
calling it brings into Names, and what Range.Formula2 shows of it. Rival rules, written before measuring. A. _xlfn.NAME comes
exactly where the file prefixes the function, IFERROR being an exception to list. B. _xlfn.SINGLE comes exactly where Formula2
shows an @.

    python scripts/measure_xlfn_functions.py

writes tests/fixtures/xlfn_functions.json, which tests/test_excel_xlfn_names.py uses.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
FUNCTIONS: dict[str, dict[str, str]] = json.loads((ROOT / "tests/fixtures/formula_prefixes/prefixes.json").read_text(encoding="utf-8"))["functions"]
BATCH = 40
PROBE = '''Private Function Listed(wb As Object) As String
    Dim n As Object, s As String
    For Each n In wb.Names
        s = s & "," & n.Name
    Next
    Listed = s
End Function
Function Go() As String
    Dim wb As Workbook, ws As Worksheet, c As Range, out As String, i As Long
    Application.DisplayAlerts = False
    Dim calls({last}) As String
{calls}
    For i = 0 To UBound(calls)
        Set wb = Workbooks.Add(-4167): Set ws = wb.Worksheets(1)
        ws.Range("A1:A4").Value = 1
        Set c = ws.Range("H9")
        On Error Resume Next
        c.Formula = calls(i)
        If Err.Number <> 0 Then
            out = out & vbTab & "E" & Err.Number
            Err.Clear
        Else
            out = out & vbTab & c.Formula2 & "~" & Listed(wb)
        End If
        On Error GoTo 0
        wb.Close False
    Next
    Go = Mid$(out, 2)
End Function
'''


def main() -> None:
    from pyvbaharness import ExcelSession, HarnessConfig

    names = sorted(FUNCTIONS)
    rows: dict[str, str] = {}
    for at in range(0, len(names), BATCH):
        part = names[at:at + BATCH]
        calls = "\n".join(f'    calls({index}) = "=' + FUNCTIONS[name]["call"].replace('"', '""') + '"' for index, name in enumerate(part))
        source = PROBE.format(calls=calls, last=len(part) - 1)
        with ExcelSession(HarnessConfig(lock_wait_s=900.0)) as excel:
            excel.new_document()
            result = excel.run_vba(source, "Go", (), timeout=120.0)
        answers = str(result.value).split("\t") if result.ok else [f"failed {result.outcome}"] * len(part)
        for name, answer in zip(part, answers):
            rows[name] = answer
        print(at, len(rows), flush=True)
    (ROOT / "tests/fixtures/xlfn_functions.json").write_text(json.dumps({"measured_at": datetime.now(timezone.utc).isoformat(), "probe": PROBE, "rows": rows}, indent=1) + "\n", encoding="utf-8")


if __name__ == "__main__":
    main()
