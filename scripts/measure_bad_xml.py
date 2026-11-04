"""What Workbooks.Open does with a package part that is not well-formed XML. A book Excel saved, holding a pivot table, a chart
and a name, has each part broken in turn, by a reference to a character XML has no room for (a lone surrogate, as a writer
that splits an emoji makes) or by a cut-off tag. Rival rules, written before measuring: A. Excel refuses every one with error
1004. B. Excel repairs what it can do without (a pivot table, a chart), opening the rest, with alerts off.

    python scripts/measure_bad_xml.py

writes tests/fixtures/bad_xml/, which tests/test_excel_bad_xml.py replays.
"""
from __future__ import annotations

import json
import re
import shutil
import sys
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))
OUT = ROOT / "tests/fixtures/bad_xml"
MAKE = '''Sub Make(ByVal file As String)
    Dim wb As Workbook, ws As Worksheet, s As Worksheet, p As Object, o As Object
    Application.DisplayAlerts = False
    Set wb = Workbooks.Add(-4167): Set ws = wb.Worksheets(1)
    Set s = wb.Worksheets.Add(After:=ws)
    ws.Range("A1:B1").Value = Array("k", "v")
    ws.Range("A2").Value = "a": ws.Range("B2").Value = 1
    ws.Range("A3").Value = "b": ws.Range("B3").Value = 2
    wb.Names.Add "Far", "=Sheet1!$A$1"
    Set p = wb.PivotCaches.Create(1, ws.Range("A1:B3")).CreatePivotTable(s.Range("A3"), "P")
    p.PivotFields("k").Orientation = 1
    p.AddDataField p.PivotFields("v"), "Sum v", -4157
    Set o = ws.ChartObjects.Add(200, 10, 200, 100): o.Chart.SetSourceData ws.Range("A1:B3")
    ws.Activate
    wb.SaveAs file, 51
    wb.Close False
End Sub
'''
READ = '''Function ReadIt(ByVal file As String) As String
    Dim wb As Workbook, s As Object, x As String
    Application.DisplayAlerts = False
    On Error Resume Next
    Set wb = Workbooks.Open(file)
    x = Err.Number & "|" & Err.Description
    Err.Clear
    If Not wb Is Nothing Then
        x = x & "|sheets " & wb.Worksheets.Count & "|a2 " & wb.Worksheets(1).Range("A2").Value & "|charts " & wb.Worksheets(1).ChartObjects.Count
        x = x & "|pivots " & wb.Worksheets(2).PivotTables.Count & "|names " & wb.Names.Count
        wb.Close False
    End If
    ReadIt = x
End Function
'''
BROKEN = {"surrogate": "&#55357;", "cut_off": "<"}
#: The parts, each as a pattern for its name in the package.
PARTS = {"workbook": r"xl/workbook\.xml", "sheet": r"xl/worksheets/sheet1\.xml", "styles": r"xl/styles\.xml", "strings": r"xl/sharedStrings\.xml",
         "theme": r"xl/theme/theme1\.xml", "drawing": r"xl/drawings/drawing1\.xml", "chart": r"xl/charts/chart1\.xml", "pivot": r"xl/pivotTables/pivotTable1\.xml",
         "pivot_cache": r"xl/pivotCache/pivotCacheDefinition1\.xml", "core": r"docProps/core\.xml", "app": r"docProps/app\.xml",
         "content_types": r"\[Content_Types\]\.xml", "book_rels": r"xl/_rels/workbook\.xml\.rels"}


def broken(source: Path, target: Path, part: str, how: str) -> bool:
    """A copy with a part broken just before its closing tag, as text inside it; False where the book has no such part."""
    hit = False
    with zipfile.ZipFile(source) as given, zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as made:
        for info in given.infolist():
            data = given.read(info.filename)
            if re.fullmatch(part, info.filename):
                at = data.rfind(b"</")
                data = data[:at] + BROKEN[how].encode() + data[at:]
                hit = True
            made.writestr(info, data)
    return hit


def main() -> None:
    from fixture_workbook import strip_save_path
    from pyvbaharness import ExcelSession, HarnessConfig

    OUT.mkdir(parents=True, exist_ok=True)
    folder = Path(tempfile.mkdtemp(prefix="pyopenvba_badxml_"))
    rows: list[dict[str, str]] = []
    try:
        base = folder / "base.xlsx"
        with ExcelSession(HarnessConfig(lock_wait_s=900.0)) as excel:
            excel.new_document()
            made = excel.run_vba(MAKE, "Make", (str(base),), timeout=60.0)
            assert made.ok, made.message
            shutil.copy(base, OUT / "base.xlsx")
            strip_save_path(OUT / "base.xlsx")
            plain = excel.run_vba(READ, "ReadIt", (str(base),), timeout=60.0, line_numbers=False)
            rows.append({"part": "", "how": "", "answer": str(plain.value)})
            print("unbroken", plain.value, flush=True)
        for label, part in PARTS.items():
            for how in BROKEN:
                file = folder / f"{label}_{how}.xlsx"
                if not broken(OUT / "base.xlsx", file, part, how):
                    print(label, "no part", flush=True)
                    continue
                with ExcelSession(HarnessConfig(lock_wait_s=900.0)) as excel:
                    excel.new_document()
                    read = excel.run_vba(READ, "ReadIt", (str(file),), timeout=60.0, line_numbers=False)
                answer = str(read.value) if read.ok else f"failed {read.outcome} {read.message[:80]}"
                rows.append({"part": label, "how": how, "answer": answer})
                print(label, how, answer, flush=True)
    finally:
        shutil.rmtree(folder, ignore_errors=True)
    (OUT / "bad_xml.json").write_text(json.dumps({"measured_at": datetime.now(timezone.utc).isoformat(), "read": READ, "parts": PARTS, "broken": BROKEN, "rows": rows}, indent=1) + "\n",
                                      encoding="utf-8")


if __name__ == "__main__":
    main()
