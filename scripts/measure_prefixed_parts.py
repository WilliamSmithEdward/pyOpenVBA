"""How Excel reads and saves a workbook whose SpreadsheetML parts spell their namespace under a prefix.

A writer may spell a part's elements as ``<x:worksheet xmlns:x="...">``
where Excel writes ``<worksheet xmlns="...">``; to a reader that follows
namespaces the two are the same XML. This copies six workbooks Excel saved
for other measurements, each part in SpreadsheetML's namespace spelled
under the prefix ``x``, and one more with its stylesheet alone so spelled,
and has Excel open each copy and the workbook it was made from, read every
cell's formula, number format, bold and fill, each sheet's conditional
formats and tables, and save the copy again.

    python scripts/measure_prefixed_parts.py

writes tests/fixtures/prefixed_parts/: each copy, Excel's save of it, and
prefixed_parts.json, which tests/test_excel_prefixed_parts.py replays.
"""

from __future__ import annotations

import json
import re
import shutil
import sys
import zipfile
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "scripts"))
from fixture_workbook import strip_save_path  # noqa: E402

FIXTURES = ROOT / "tests" / "fixtures"
OUT = FIXTURES / "prefixed_parts"
MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
#: Each copy's name, and the workbook it is made from: styles, tables, pivot tables, conditional formats, a sort
#: and a filter; xfs of every kind; formulas and their calculation chain; notes and links; a query's connection
#: and query table; and dynamic arrays' metadata.
SOURCES = {
    "pivots": "dxf_table/order.xlsx",
    "styles": "unchanged_save/base.xlsx",
    "formulas": "formula_storage/delete_row.xlsx",
    "notes": "hyperlinks/rels.xlsx",
    "query": "power_query/loaded_to_sheet.xlsx",
    "arrays": "at_sign.xlsx",
    "stylesheet": "unchanged_save/base.xlsx",
}
#: The copies with only some parts spelled under the prefix, and those parts.
ONLY = {"stylesheet": {"xl/styles.xml"}}

READ = '''Public Function Read(path As String, saved As String) As String
Dim wb As Object, ws As Object, c As Object, f As Object, out As String, t As Object
Application.DisplayAlerts = False
On Error Resume Next
Set wb = Workbooks.Open(path)
If wb Is Nothing Then
    Read = "E" & Err.Number
    Exit Function
End If
For Each ws In wb.Worksheets
    out = out & ws.Name & ":"
    For Each c In ws.UsedRange.Cells
        If Not IsEmpty(c.Value) Then out = out & c.Address(False, False) & "=" & CStr(c.Formula) & "/" & _
            c.NumberFormat & "/" & c.Font.Bold & "/" & c.Interior.Color & ","
    Next
    For Each f In ws.Cells.FormatConditions
        out = out & "rule " & f.AppliesTo.Address(False, False) & "@" & f.Priority & "=" & f.Interior.Color & ","
    Next
    For Each t In ws.ListObjects
        out = out & "table " & t.Name & " " & t.Range.Address(False, False) & ","
    Next
    out = out & ";"
Next
If Len(saved) > 0 Then wb.SaveAs saved, 51
wb.Close False
Read = out
End Function
'''


def prefixed(text: str) -> str:
    """A part whose root declares SpreadsheetML's namespace as its default, with that namespace under ``x``."""
    start = re.search(r"<(?![?!])[^>]*>", text)
    if start is None or f'xmlns="{MAIN}"' not in start.group():
        return text
    body = re.sub(r"<(/?)([A-Za-z_][\w.-]*)(?=[\s/>])", lambda tag: f"<{tag.group(1)}x:{tag.group(2)}",
                  text[start.start():])
    return text[:start.start()] + body.replace(f'xmlns="{MAIN}"', f'xmlns:x="{MAIN}"', 1)


def roots(path: Path) -> dict[str, str]:
    """Each part of SpreadsheetML's under xl/, and its root element as the part spells it."""
    out: dict[str, str] = {}
    with zipfile.ZipFile(path) as package:
        for part in sorted(package.namelist()):
            if part.startswith("xl/") and part.endswith(".xml"):
                text = package.read(part).decode("utf-8")
                start = re.search(r"<(?![?!])([\w:.-]+)[^>]*>", text)
                if start is not None and MAIN in start.group():
                    out[part] = start.group(1)
    return out


def main() -> None:
    if OUT.exists():
        shutil.rmtree(OUT)
    OUT.mkdir(parents=True)
    for name, source in SOURCES.items():
        with zipfile.ZipFile(FIXTURES / source) as package, \
                zipfile.ZipFile(OUT / f"{name}.xlsx", "w", zipfile.ZIP_DEFLATED) as out:
            for info in package.infolist():
                data = package.read(info.filename)
                if info.filename.startswith("xl/") and info.filename.endswith(".xml") \
                        and info.filename in ONLY.get(name, {info.filename}):
                    data = prefixed(data.decode("utf-8")).encode("utf-8")
                out.writestr(info, data)
    cases: list[dict[str, object]] = []
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        for name, source in SOURCES.items():
            original = excel.run_vba(READ, "Read", (str(FIXTURES / source), ""), timeout=600.0)
            assert original.ok, f"{original.outcome}: {original.message} {original.error}"
            copy = excel.run_vba(READ, "Read", (str(OUT / f"{name}.xlsx"), str(OUT / f"{name}_resaved.xlsx")),
                                 timeout=600.0)
            assert copy.ok, f"{copy.outcome}: {copy.message} {copy.error}"
            cases.append({"name": name, "source": source, "answer": str(copy.value),
                          "alike": str(copy.value) == str(original.value)})
    for case in cases:
        resaved = OUT / f"{case['name']}_resaved.xlsx"
        strip_save_path(resaved)
        case["prefixed"] = roots(OUT / f"{case['name']}.xlsx")
        case["resaved"] = roots(resaved)
    record = {"read": READ, "cases": cases}
    (OUT / "prefixed_parts.json").write_text(json.dumps(record, indent=1, ensure_ascii=False) + "\n",
                                             encoding="utf-8")
    for case in cases:
        print("==", case["name"], "read alike:", case["alike"])
        print("   Excel's save:", case["resaved"])


if __name__ == "__main__":
    main()
