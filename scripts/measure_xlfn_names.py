"""When Names holds a hidden _xlfn name. A cell written with =Nowhere through Range.Formula put _xlfn.SINGLE, =#NAME?, into
wb.Names though the file lists no such name. Rival rules, written before measuring. A. Any formula Excel keeps with an @ adds
it, from the moment it is written. B. Only an @ on a name that is not there. C. Any function a file spells _xlfn. adds its own
name too (CONCAT, TEXTJOIN). And whether it goes when the cell is cleared, comes back on reopening, counts in Names.Count,
and answers Names("_xlfn.SINGLE").

    python scripts/measure_xlfn_names.py

writes tests/fixtures/xlfn_names.json, which tests/test_excel_xlfn_names.py replays.
"""
from __future__ import annotations

import json
import re
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
#: Name: what is written, each in a new workbook whose A1:A3 hold 1, 2, 3.
CASES: dict[str, str] = {
    "nothing": "",
    "plain": 'c.Formula = "=A1+1"',
    "unknown_name": 'c.Formula = "=Nowhere"',
    "unknown_name2": 'c.Formula2 = "=Nowhere"',
    "unknown_in_sum": 'c.Formula = "=SUM(Nowhere)"',
    "unknown_plus": 'c.Formula = "=Nowhere+1"',
    "known_name": 'wb.Names.Add "Here", "=Sheet1!$A$1": c.Formula = "=Here"',
    "known_block": 'wb.Names.Add "Here", "=Sheet1!$A$1:$A$3": c.Formula = "=Here"',
    "block": 'c.Formula = "=A1:A3"',
    "block2": 'c.Formula2 = "=A1:A3"',
    "at_block": 'c.Formula2 = "=@A1:A3"',
    "at_unknown": 'c.Formula2 = "=@Nowhere"',
    "block_in_len": 'c.Formula = "=LEN(A1:A3)"',
    "udf_like": 'c.Formula = "=NoFunction(1)"',
    "concat": 'c.Formula = "=CONCAT(A1:A3)"',
    "textjoin": 'c.Formula = "=TEXTJOIN("","",TRUE,A1:A3)"',
    "xlookup": 'c.Formula = "=XLOOKUP(1,A1:A3,A1:A3)"',
    "sequence": 'c.Formula2 = "=SEQUENCE(2)"',
    "two_cells": 'c.Formula = "=Nowhere": ws.Range("D2").Formula = "=Elsewhere"',
    "cleared": 'c.Formula = "=Nowhere": c.ClearContents',
    "replaced": 'c.Formula = "=Nowhere": c.Value = 1',
    "replaced_formula": 'c.Formula = "=Nowhere": c.Formula = "=1+1"',
    "named_after": 'c.Formula = "=Nowhere": wb.Names.Add "Nowhere", "=Sheet1!$A$1"',
    "sort": 'c.Formula2 = "=SORT(A1:A3)"',
    "filter": 'c.Formula2 = "=FILTER(A1:A3,A1:A3>1)"',
    "unique": 'c.Formula2 = "=UNIQUE(A1:A3)"',
    "let": 'c.Formula2 = "=LET(x,1,x+1)"',
    "lambda": 'c.Formula2 = "=LAMBDA(x,x+1)(2)"',
    "ifs": 'c.Formula = "=IFS(A1>0,1)"',
    "maxifs": 'c.Formula = "=MAXIFS(A1:A3,A1:A3,"">0"")"',
    "stdev_s": 'c.Formula = "=STDEV.S(A1:A3)"',
    "aggregate": 'c.Formula = "=AGGREGATE(9,6,A1:A3)"',
    "iferror": 'c.Formula = "=IFERROR(A1,0)"',
    "both": 'c.Formula = "=CONCAT(A1:A3)": ws.Range("D2").Formula = "=Nowhere": ws.Range("D3").Formula = "=XLOOKUP(1,A1:A3,A1:A3)"',
    "name_with_function": 'wb.Names.Add "Calc", "=CONCAT(Sheet1!$A$1:$A$3)"',
    "name_at": 'wb.Names.Add "Calc", "=Nowhere"',
    "cond_format": 'ws.Range("E1").FormatConditions.Add 2, , "=CONCAT($A$1:$A$3)=""x"""',
    "with_a_name": 'wb.Names.Add "Zed", "=Sheet1!$A$1": wb.Names.Add "Abc", "=Sheet1!$A$2": c.Formula = "=Nowhere"',
}
PROBE = '''Private Function Listed(wb As Object) As String
    Dim n As Object, s As String, i As Long
    s = wb.Names.Count & ":"
    For Each n In wb.Names
        s = s & " [" & n.Name & "=" & n.RefersTo & " vis " & n.Visible & " idx " & n.Index & "]"
    Next
    Listed = s
End Function
Function Go(ByVal file As String) As String
    Dim wb As Workbook, ws As Worksheet, c As Range, t As String, v As Variant
    Application.DisplayAlerts = False
    Set wb = Workbooks.Add(-4167): Set ws = wb.Worksheets(1)
    ws.Range("A1").Value = 1: ws.Range("A2").Value = 2: ws.Range("A3").Value = 3
    Set c = ws.Range("D1")
    {action}
    t = "formula " & c.Formula & " text " & c.Text & " | " & Listed(wb)
    On Error Resume Next
    v = wb.Names("_xlfn.SINGLE").Name
    t = t & " | item E" & Err.Number
    Err.Clear
    On Error GoTo 0
    wb.SaveAs file, 51
    t = t & " | saved " & Listed(wb)
    wb.Close False
    Set wb = Workbooks.Open(file)
    t = t & " | reopen " & Listed(wb) & " | " & wb.Worksheets(1).Range("D1").Formula
    wb.Worksheets(1).Range("D1").ClearContents
    t = t & " | cleared " & Listed(wb)
    wb.Close False
    Go = t
End Function
'''


def _spelled(file: Path) -> str:
    """The cell formulas and defined names as Excel's file has them."""
    if not file.exists():
        return ""
    with zipfile.ZipFile(file) as package:
        sheet = package.read("xl/worksheets/sheet1.xml").decode("utf-8")
        book = package.read("xl/workbook.xml").decode("utf-8")
    return " ".join(re.findall(r"<f\b[^>]*>.*?</f>|<c\b[^>]*\bcm=\"1\"[^>]*>", sheet) + re.findall(r"<definedName\b.*?</definedName>", book))


def main() -> None:
    from pyvbaharness import ExcelSession, HarnessConfig

    folder = Path(tempfile.mkdtemp(prefix="pyopenvba_xlfn_"))
    rows: list[dict[str, str]] = []
    for name, action in CASES.items():
        file = folder / f"{name}.xlsx"
        with ExcelSession(HarnessConfig(lock_wait_s=900.0)) as excel:
            excel.new_document()
            result = excel.run_vba(PROBE.format(action=action), "Go", (str(file),), timeout=60.0, line_numbers=False)
        answer = str(result.value) if result.ok else f"failed {result.outcome} {result.message[:80]}"
        rows.append({"name": name, "answer": answer.replace(file.name, "<book>"), "file": _spelled(file)})
        print(name, "|", rows[-1]["file"], "|", rows[-1]["answer"], flush=True)
    (ROOT / "tests/fixtures/xlfn_names.json").write_text(json.dumps({"measured_at": datetime.now(timezone.utc).isoformat(), "probe": PROBE, "cases": CASES, "rows": rows}, indent=1) + "\n",
                                                         encoding="utf-8")


if __name__ == "__main__":
    main()
