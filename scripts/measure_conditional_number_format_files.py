"""How Excel reads a rule's number format that a file numbers or spells another way than Excel writes it.

Excel writes a rule's number format with the id its own list gives a
built-in one, a custom one numbered with the cells' formats, and every
literal escaped. Other programs write what they like. This makes a
workbook in Excel with two sheets, a rule on each and two custom formats
on cells, then writes copies of it whose rules' numFmts take other ids
and spellings: a built-in code under a custom id, a built-in id over
another code, the ids cells give built-in formats, a cell's custom id,
a gap, one code twice, General, and spellings Excel never writes. Excel
opens each copy, reads each rule's NumberFormat and what its cell shows,
and saves it again.

    python scripts/measure_conditional_number_format_files.py

writes tests/fixtures/conditional_number_formats/foreign/: each copy and
Excel's save of it, and foreign.json beside the folder, which
tests/test_excel_conditional_number_formats.py replays.
"""

from __future__ import annotations

import json
import re
import shutil
import sys
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape

from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "scripts"))
from fixture_workbook import strip_save_path  # noqa: E402
from measure_conditional_number_formats import READER, number_formats  # noqa: E402

OUT = ROOT / "tests" / "fixtures" / "conditional_number_formats"
FOLDER = OUT / "foreign"
#: The cells' own formats every copy holds: A1's and A2's, the second's id out of order.
CELLS = ((164, "0.0"), (190, "0.0000"))
#: name -> the numFmt the rule on s1 and the rule on s2 hold, each as (numFmtId, formatCode as the file spells it).
CASES: dict[str, tuple[tuple[int, str], tuple[int, str]]] = {
    # A built-in code under a custom id, and a built-in id over another code.
    "listed_code_custom_id": ((170, "0%"), (181, "0.0%")),
    "listed_id_other_code": ((13, "0.0"), (181, "0.0%")),
    # Built-in formats by the ids cells number them with.
    "cell_builtin_ids": ((9, "0%"), (40, "#,##0.00_);[Red](#,##0.00)")),
    "cell_builtin_currency": ((5, '"$"#,##0_);\\("$"#,##0\\)'), (7, "$#,##0.00_);($#,##0.00)")),
    # The custom id a cell's format has, over another code and over that format's own code.
    "cell_id_other_code": ((164, "0.00000"), (181, "0.0%")),
    "cell_id_cell_code": ((190, "0.0000"), (165, "0.0")),
    # A gap in the ids and one code under two of them, and an id below 164 no built-in format has.
    "gap_and_duplicate": ((300, "0.000"), (171, "0.000")),
    "low_id_cell_code": ((100, "0.0"), (181, "0.0%")),
    # General under a custom id, and spellings Excel does not write.
    "general_custom_id": ((172, "General"), (181, "0.0%")),
    "foreign_spelling": ((170, '#,##0 "kr"'), (181, "0.0 (x)")),
    "dollar_bare": ((170, "$0.00"), (181, "0.00$")),
}

BUILD = '''Public Function Build() As String
Dim wb As Object, ws As Object, f As Object, name As Variant
Application.DisplayAlerts = False
Set wb = Workbooks.Add(xlWBATWorksheet)
For Each name In Array("s1", "s2")
    If name = "s1" Then
        Set ws = wb.Worksheets(1)
        ws.Range("A1").Value = 1.5
        ws.Range("A1").NumberFormat = "0.0"
        ws.Range("A2").Value = 2.5
        ws.Range("A2").NumberFormat = "0.0000"
    Else
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(1))
    End If
    ws.Name = name
    ws.Columns("B").ColumnWidth = 24
    ws.Range("B2:B3").Value = 50
    Set f = ws.Range("B2:B3").FormatConditions.Add(xlCellValue, xlGreater, "22")
    f.Interior.Color = RGB(255, 0, 0)
    f.NumberFormat = IIf(name = "s1", "0.000", "0.00000")
Next
wb.SaveAs "{path}", 51
wb.Close False
Build = "ok"
End Function
'''


def read_code(names: list[str]) -> str:
    lines = [READER, "Public Function Read() As String", "Dim wb As Object, out As String, name As Variant",
             "Application.DisplayAlerts = False", f"For Each name In Array({', '.join(repr(one) for one in names)})"
             .replace("'", '"'),
             "    On Error Resume Next", "    Err.Clear", "    Set wb = Nothing",
             f'    Set wb = Workbooks.Open("{FOLDER}\\" & name & ".xlsx")',
             '    If wb Is Nothing Then',
             '        out = out & name & "^E" & Err.Number & "|"',
             "    Else",
             '        out = out & name & "^" & Shown(wb.Worksheets("s1")) & "~" & wb.Worksheets("s1").Range("B2").Text'
             ' & "~" & Shown(wb.Worksheets("s2")) & "~" & wb.Worksheets("s2").Range("B2").Text'
             ' & "~" & wb.Worksheets("s1").Range("A1").NumberFormat & "~" & wb.Worksheets("s1").Range("A2").NumberFormat'
             ' & "|"',
             f'        wb.SaveAs "{FOLDER}\\" & name & "_resaved.xlsx", 51',
             "        wb.Close False",
             "    End If",
             "    On Error GoTo 0",
             "Next", "Read = out", "End Function"]
    return "\n".join(lines) + "\n"


def _sheet_rules(package: zipfile.ZipFile) -> dict[str, int]:
    """Each sheet's rule's dxfId, by the sheet's name."""
    workbook = package.read("xl/workbook.xml").decode("utf-8")
    targets = {found.group(1): found.group(2) for found in re.finditer(
        r'<Relationship\b[^>]*?Id="([^"]+)"[^>]*?Target="([^"]+)"', package.read("xl/_rels/workbook.xml.rels").decode())}
    out: dict[str, int] = {}
    for name, rid in re.findall(r'<sheet\b[^>]*?name="([^"]+)"[^>]*?r:id="([^"]+)"', workbook):
        target = targets[rid].lstrip("/")
        text = package.read(target if target.startswith("xl/") else "xl/" + target).decode("utf-8")
        rule = re.search(r'<cfRule\b[^>]*?dxfId="(\d+)"', text)
        assert rule is not None
        out[name] = int(rule.group(1))
    return out


def crafted(base: Path, target: Path, rules: tuple[tuple[int, str], tuple[int, str]]) -> None:
    """``base`` with the cells' formats numbered as CELLS says and each rule's numFmt the one ``rules`` gives."""
    with zipfile.ZipFile(base) as package:
        styles = package.read("xl/styles.xml").decode("utf-8")
        at = _sheet_rules(package)
        listed = re.search(r"<numFmts\b.*?</numFmts>", styles, re.DOTALL)
        assert listed is not None
        given = dict(re.findall(r'<numFmt numFmtId="(\d+)" formatCode="([^"]*)"/>', listed.group()))
        # Excel gave A1's format 164 and A2's 165; A2's takes 190 here.
        assert given == {"164": "0.0", "165": "0.0000"}, given
        numbers = "".join(f'<numFmt numFmtId="{identifier}" formatCode="{escape(code, {chr(34): "&quot;"})}"/>'
                          for identifier, code in CELLS)
        styles = styles[:listed.start()] + f'<numFmts count="{len(CELLS)}">{numbers}</numFmts>' + styles[listed.end():]
        xfs = re.search(r"<cellXfs\b.*?</cellXfs>", styles, re.DOTALL)
        assert xfs is not None
        styles = styles[:xfs.start()] + xfs.group().replace('numFmtId="165"', 'numFmtId="190"') + styles[xfs.end():]
        dxfs = re.search(r"<dxfs\b.*?</dxfs>", styles, re.DOTALL)
        assert dxfs is not None
        entries = re.findall(r"<dxf>.*?</dxf>", dxfs.group(), re.DOTALL)
        for sheet, (identifier, code) in zip(("s1", "s2"), rules, strict=True):
            entries[at[sheet]] = re.sub(r"<numFmt [^>]*/>", lambda _: f'<numFmt numFmtId="{identifier}" formatCode='
                                        f'"{escape(code, {chr(34): "&quot;"})}"/>', entries[at[sheet]], count=1)
        opened = dxfs.group()[:dxfs.group().index(">") + 1]
        styles = styles[:dxfs.start()] + opened + "".join(entries) + "</dxfs>" + styles[dxfs.end():]
        with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as out:
            for entry in package.infolist():
                data = styles.encode("utf-8") if entry.filename == "xl/styles.xml" else package.read(entry.filename)
                out.writestr(entry, data)


def main() -> None:
    if FOLDER.exists():
        shutil.rmtree(FOLDER)
    FOLDER.mkdir(parents=True)
    base = FOLDER / "base.xlsx"
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(BUILD.replace("{path}", str(base)), "Build", timeout=300.0)
        assert result.ok and result.value == "ok", f"{result.outcome}: {result.message} {result.error}"
    strip_save_path(base)
    for name, rules in CASES.items():
        crafted(base, FOLDER / f"{name}.xlsx", rules)
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(read_code(list(CASES)), "Read", timeout=600.0)
        assert result.ok, f"{result.outcome}: {result.message} {result.error}"
    answers = dict(one.split("^", 1) for one in str(result.value).split("|") if "^" in one)
    record: dict[str, object] = {"cells": [list(one) for one in CELLS], "cases": []}
    cases: list[dict[str, object]] = []
    for name, rules in CASES.items():
        answer = answers.get(name, "")
        resaved = FOLDER / f"{name}_resaved.xlsx"
        files: dict[str, str] = {}
        cells: list[str] = []
        if resaved.exists():
            strip_save_path(resaved)
            with zipfile.ZipFile(resaved) as package:
                files, cells = number_formats(package)
        parts = answer.split("~")
        cases.append({"name": name, "rules": [list(one) for one in rules], "answer": answer,
                      "shown": parts[0::2][:2] if len(parts) == 6 else [], "text": parts[1:4:2] if len(parts) == 6 else [],
                      "cells_shown": parts[4:6] if len(parts) == 6 else [], "resaved": files, "resaved_cells": cells})
    record["cases"] = cases
    (OUT / "foreign.json").write_text(json.dumps(record, indent=1, ensure_ascii=False) + "\n", encoding="utf-8")
    for one in cases:
        print(f"{one['name']:24} {one['rules']}")
        print(f"    answer {one['answer']}")
        print(f"    resaved {one['resaved']} cells {one['resaved_cells']}")


if __name__ == "__main__":
    main()
