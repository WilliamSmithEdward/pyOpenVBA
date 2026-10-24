"""What Excel does with a workbook's own custom number formats when it opens a file and saves it unchanged.

Excel numbers a file's custom formats from 164 as it opens it, whatever
ids the file gave them, and writes those ids when it saves. This makes a
workbook in Excel whose custom formats came after a conditional format's,
so they sit at ids a session gave them: four on cells, a pivot table's
data field, and a rule kept in the worksheet's Excel 2010 extension. It
then writes copies of it numbering and spelling the cells' formats
otherwise: listed out of order, with gaps, with one nothing uses, one only
an unused xf uses, one code twice, a built-in id over another code, an id
below 164, a spelling Excel does not write, an xf naming a format the file
lacks, and the pivot's format at an id of its own. Excel opens each,
reads every cell's NumberFormat, the pivot's and the rules', and saves it
again with nothing changed.

    python scripts/measure_number_format_resave.py

writes tests/fixtures/number_format_resave/: the workbook, each copy,
Excel's save of each, and number_format_resave.json, which
tests/test_excel_number_format_resave.py replays.
"""

from __future__ import annotations

import json
import re
import shutil
import sys
import zipfile
from collections.abc import Callable
from pathlib import Path
from xml.sax.saxutils import escape

from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "scripts"))
from fixture_workbook import strip_save_path  # noqa: E402

OUT = ROOT / "tests" / "fixtures" / "number_format_resave"

BUILD = '''Public Function Build() As String
Dim wb As Object, ws As Object, data As Object, pv As Object, f As Object, pt As Object, r As Long
Application.DisplayAlerts = False
Set wb = Workbooks.Add(xlWBATWorksheet)
Set ws = wb.Worksheets(1)
ws.Name = "cells"
ws.Columns("A:C").ColumnWidth = 16
ws.Range("B2:C3").Value = 50
Set f = ws.Range("B2:B3").FormatConditions.Add(xlCellValue, xlGreater, "22")
f.Interior.Color = RGB(255, 0, 0)
f.NumberFormat = "0.00000"
ws.Range("A1:A4").Value = 1.5
ws.Range("A1").NumberFormat = "0.0"
ws.Range("A2").NumberFormat = "0.0000"
ws.Range("A3").NumberFormat = "#,##0.000"
ws.Range("A4").NumberFormat = "0.000%"
Set data = wb.Worksheets.Add(After:=ws)
data.Name = "data"
data.Range("A1").Value = "Name"
data.Range("B1").Value = "Amount"
For r = 2 To 5
    data.Cells(r, 1).Value = Chr(95 + r)
    data.Cells(r, 2).Value = r - 1
Next
Set pv = wb.Worksheets.Add(After:=data)
pv.Name = "pivot"
Set pt = wb.PivotCaches.Create(xlDatabase, "data!R1C1:R5C2").CreatePivotTable(pv.Range("A3"), "Totals")
pt.PivotFields("Name").Orientation = xlRowField
pt.AddDataField pt.PivotFields("Amount"), "Total", xlSum
pt.DataFields(1).NumberFormat = "0.000000"
Set f = ws.Range("C2:C3").FormatConditions.Add(xlCellValue, xlGreater, "=data!$B$2")
f.Interior.Color = RGB(0, 255, 0)
f.NumberFormat = "0.0000000"
wb.SaveAs "{path}", 51
wb.Close False
Build = "ok"
End Function
'''

READ = '''Private Function Shown(f As Object) As String
    On Error Resume Next
    Shown = f.NumberFormat
    If Err.Number <> 0 Then Shown = "E" & Err.Number
End Function

Public Function Read() As String
Dim wb As Object, ws As Object, out As String, name As Variant, r As Long
Application.DisplayAlerts = False
For Each name In Array({names})
    On Error Resume Next
    Err.Clear
    Set wb = Nothing
    Set wb = Workbooks.Open("{folder}\\" & name & ".xlsx")
    If wb Is Nothing Then
        out = out & name & "^E" & Err.Number & "|"
    Else
        On Error GoTo 0
        Set ws = wb.Worksheets("cells")
        out = out & name & "^"
        For r = 1 To 4
            out = out & ws.Cells(r, 1).NumberFormat & "~" & ws.Cells(r, 1).Text & "~"
        Next
        out = out & Shown(wb.Worksheets("pivot").PivotTables(1).DataFields(1)) & "~" & _
            Shown(ws.Range("B2").FormatConditions(1)) & "~" & Shown(ws.Range("C2").FormatConditions(1)) & "|"
        wb.SaveAs "{folder}\\" & name & "_resaved.xlsx", 51
        wb.Close False
    End If
    On Error GoTo 0
Next
Read = out
End Function
'''


def _part(package: zipfile.ZipFile, sheet: str) -> str:
    """The part a sheet of ``package`` is kept in."""
    workbook = package.read("xl/workbook.xml").decode("utf-8")
    targets = {found.group(1): found.group(2) for found in re.finditer(
        r'<Relationship\b[^>]*?Id="([^"]+)"[^>]*?Target="([^"]+)"', package.read("xl/_rels/workbook.xml.rels").decode())}
    rid = re.search(rf'<sheet\b[^>]*?name="{sheet}"[^>]*?r:id="([^"]+)"', workbook)
    assert rid is not None
    target = targets[rid.group(1)].lstrip("/")
    return target if target.startswith("xl/") else "xl/" + target


def numbers(path: Path) -> dict[str, object]:
    """What a file holds of its number formats: the stylesheet's numFmts, each of the sheet "cells" A1:A4's
    numFmtId, the dxfs' numFmts, the rule in the sheet's extension's numFmt, and every numFmtId in the pivot's
    parts."""
    with zipfile.ZipFile(path) as package:
        styles = package.read("xl/styles.xml").decode("utf-8")
        sheet = package.read(_part(package, "cells")).decode("utf-8")
        pivots = {name: re.findall(r'<(\w+)\b[^>]*?\bnumFmtId="(\d+)"', package.read(name).decode("utf-8"))
                  for name in package.namelist() if name.startswith("xl/pivot") and name.endswith(".xml")}
    listed = re.search(r"<numFmts\b.*?</numFmts>", styles, re.DOTALL)
    cell_xfs = re.search(r"<cellXfs\b.*?</cellXfs>", styles, re.DOTALL)
    assert cell_xfs is not None
    named = [found.group(1) if found else "" for found in
             (re.search(r'numFmtId="(\d+)"', xf) for xf in re.findall(r"<xf\b[^>]*?>", cell_xfs.group()))]
    cells = {found.group(1): named[int(found.group(2))] for found in re.finditer(r'<c r="(A[1-4])" s="(\d+)"', sheet)}
    dxfs = re.search(r"<dxfs\b.*?</dxfs>", styles, re.DOTALL)
    extension = re.search(r"<x14:dxf>.*?</x14:dxf>", sheet, re.DOTALL)
    return {"numFmts": re.findall(r"<numFmt [^>]*/>", listed.group()) if listed else [],
            "xfs": named,
            "cells": cells,
            "dxfs": re.findall(r"<numFmt [^>]*/>", dxfs.group()) if dxfs else [],
            "extension": re.findall(r"<numFmt [^>]*/>", extension.group()) if extension else [],
            "pivot": {name: [f"{tag}:{identifier}" for tag, identifier in found] for name, found in pivots.items()}}


def _cell_xf(sheet: str, cell: str) -> int:
    found = re.search(rf'<c r="{cell}" s="(\d+)"', sheet)
    assert found is not None
    return int(found.group(1))


def _numbered(styles: str, code: str, identifier: str) -> str:
    """``styles`` with the numFmt holding ``code`` given ``identifier``, in the numFmts and in every cell xf."""
    found = re.search(rf'<numFmt numFmtId="(\d+)" formatCode="{re.escape(escape(code, {chr(34): "&quot;"}))}"/>',
                      styles)
    assert found is not None, code
    old = found.group(1)
    styles = styles.replace(found.group(), found.group().replace(f'numFmtId="{old}"', f'numFmtId="{identifier}"'))
    xfs = re.search(r"<cellXfs\b.*?</cellXfs>", styles, re.DOTALL)
    assert xfs is not None
    return styles[:xfs.start()] + xfs.group().replace(f'numFmtId="{old}"', f'numFmtId="{identifier}"') \
        + styles[xfs.end():]


def _listing(styles: str, change: Callable[[list[str]], list[str]]) -> str:
    """``styles`` with its numFmts as ``change`` makes them."""
    listed = re.search(r"<numFmts\b[^>]*>(.*?)</numFmts>", styles, re.DOTALL)
    assert listed is not None
    entries = change(re.findall(r"<numFmt [^>]*/>", listed.group(1)))
    return styles[:listed.start()] + f'<numFmts count="{len(entries)}">' + "".join(entries) + "</numFmts>" \
        + styles[listed.end():]


def _xf_of(styles: str, sheet: str, cell: str, identifier: str) -> str:
    """``styles`` with the cell xf ``cell`` points at naming the format ``identifier``."""
    index = _cell_xf(sheet, cell)
    xfs = re.search(r"<cellXfs\b[^>]*>(.*?)</cellXfs>", styles, re.DOTALL)
    assert xfs is not None
    entries = re.findall(r"<xf\b[^>]*?(?:/>|>.*?</xf>)", xfs.group(1), re.DOTALL)
    entries[index] = re.sub(r'numFmtId="\d+"', f'numFmtId="{identifier}"', entries[index], count=1)
    return styles[:xfs.start(1)] + "".join(entries) + styles[xfs.end(1):]


def _with_xf(styles: str, identifier: str) -> str:
    """``styles`` with one more cell xf, naming the format ``identifier``, that no cell points at."""
    xfs = re.search(r'<cellXfs count="(\d+)">(.*?)</cellXfs>', styles, re.DOTALL)
    assert xfs is not None
    added = f'<xf numFmtId="{identifier}" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
    return styles[:xfs.start()] + f'<cellXfs count="{int(xfs.group(1)) + 1}">' + xfs.group(2) + added \
        + "</cellXfs>" + styles[xfs.end():]


def variants(styles: str, sheet: str) -> dict[str, tuple[str, dict[str, str]]]:
    """Each copy: its stylesheet, and any pivot part's text changes as (old, new) by part."""
    entry = lambda identifier, code: (  # noqa: E731
        f'<numFmt numFmtId="{identifier}" formatCode="{escape(code, {chr(34): "&quot;"})}"/>')
    return {
        "base": (styles, {}),
        "out_of_order": (_listing(styles, lambda entries: entries[::-1]), {}),
        "gaps": (_numbered(_numbered(_numbered(_numbered(styles, "0.0", "200"), "0.0000", "300"), "#,##0.000",
                                     "250"), "0.000%", "400"), {}),
        "unused": (_listing(styles, lambda entries: [entry(180, "0.0E+0"), *entries]), {}),
        "unused_xf": (_with_xf(_listing(styles, lambda entries: [*entries, entry(181, "0.0_)")]), "181"), {}),
        "duplicate": (_xf_of(_listing(styles, lambda entries: [*entries, entry(171, "0.0")]), sheet, "A4", "171"), {}),
        "builtin_id": (_xf_of(_listing(styles, lambda entries: [*entries, entry(2, "0.000")]), sheet, "A3", "2"), {}),
        "low_id": (_numbered(styles, "0.0000", "100"), {}),
        "spelling": (styles.replace('formatCode="0.0"/>', f'formatCode="{escape("#,##0 kr")}"/>', 1), {}),
        "missing": (_xf_of(styles, sheet, "A4", "190"), {}),
        "pivot_id": (_numbered(styles, "0.000000", "500"), {"pivot": "500"}),
    }


def write(base: Path, target: Path, styles: str, pivot: str | None) -> None:
    """``base`` with the stylesheet ``styles``, and the pivot's own format, 169, named ``pivot`` where given."""
    with zipfile.ZipFile(base) as package, zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as out:
        for info in package.infolist():
            data = package.read(info.filename)
            if info.filename == "xl/styles.xml":
                data = styles.encode("utf-8")
            elif pivot is not None and info.filename.startswith("xl/pivot") and info.filename.endswith(".xml"):
                data = data.replace(b'numFmtId="169"', f'numFmtId="{pivot}"'.encode())
            out.writestr(info, data)


def main() -> None:
    if OUT.exists():
        shutil.rmtree(OUT)
    OUT.mkdir(parents=True)
    base = OUT / "base.xlsx"
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(BUILD.replace("{path}", str(base)), "Build", timeout=300.0)
        assert result.ok and result.value == "ok", f"{result.outcome}: {result.message} {result.error}"
    strip_save_path(base)
    with zipfile.ZipFile(base) as package:
        styles = package.read("xl/styles.xml").decode("utf-8")
        sheet = package.read(_part(package, "cells")).decode("utf-8")
    made = variants(styles, sheet)
    for name, (text, changes) in made.items():
        if name != "base":
            write(base, OUT / f"{name}.xlsx", text, changes.get("pivot"))
    names = list(made)
    code = READ.replace("{names}", ", ".join(f'"{name}"' for name in names)).replace("{folder}", str(OUT))
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(code, "Read", timeout=600.0)
        assert result.ok, f"{result.outcome}: {result.message} {result.error}"
    answers = dict(one.split("^", 1) for one in str(result.value).split("|") if "^" in one)
    cases: list[dict[str, object]] = []
    for name in names:
        resaved = OUT / f"{name}_resaved.xlsx"
        if resaved.exists():
            strip_save_path(resaved)
        cases.append({"name": name, "answer": answers.get(name, ""), "file": numbers(OUT / f"{name}.xlsx"),
                      "resaved": numbers(resaved) if resaved.exists() else {}})
    (OUT / "number_format_resave.json").write_text(json.dumps({"read": READ, "cases": cases}, indent=1,
                                                              ensure_ascii=False) + "\n", encoding="utf-8")
    for one in cases:
        print("==", one["name"], ascii(one["answer"]))
        print("   file   ", ascii(one["file"]))
        print("   resaved", ascii(one["resaved"]))


if __name__ == "__main__":
    main()
