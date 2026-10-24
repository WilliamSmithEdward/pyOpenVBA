"""The number format a conditional format shows: what it reads back, and what a file keeps of it.

Each layout sits on a sheet of its own in one new workbook and sets one
rule's NumberFormat to a code, reading it back at once: Excel's built-in
formats in the order its Custom list shows them, codes Excel spells
another way, and custom ones, some of them also on a cell. The workbook
is saved, opened again, each rule's NumberFormat read back, and saved
once more; last, a rule made in the opened workbook takes a new custom
format, to show the id Excel gives it then.

    python scripts/measure_conditional_number_formats.py

Every printable character, bare and escaped, after a digit, before one,
between date codes and after the text placeholder, is then a rule's
format and a cell's in workbooks of their own, a hundred and twenty to
each, since a workbook holds about 250 custom formats.

writes tests/fixtures/conditional_number_formats/: the workbook as Excel
saved it and saved it again, chars_*.xlsx, and
conditional_number_formats.json, which
tests/test_excel_conditional_number_formats.py replays.
"""

from __future__ import annotations

import json
import re
import sys
import zipfile
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "scripts"))
from fixture_workbook import strip_save_path  # noqa: E402

OUT = ROOT / "tests" / "fixtures" / "conditional_number_formats"

#: Excel's built-in formats, as its Custom list shows them in en-US, in that list's order.
BUILT_IN = ["General", "0", "0.00", "#,##0", "#,##0.00", "#,##0_);(#,##0)", "#,##0_);[Red](#,##0)",
            "#,##0.00_);(#,##0.00)", "#,##0.00_);[Red](#,##0.00)", "$#,##0_);($#,##0)", "$#,##0_);[Red]($#,##0)",
            "$#,##0.00_);($#,##0.00)", "$#,##0.00_);[Red]($#,##0.00)", "0%", "0.00%", "0.00E+00", "##0.0E+0",
            "# ?/?", "# ??/??", "m/d/yyyy", "d-mmm-yy", "d-mmm", "mmm-yy", "h:mm AM/PM", "h:mm:ss AM/PM", "h:mm",
            "h:mm:ss", "m/d/yyyy h:mm", "mm:ss", "mm:ss.0", "@", "[h]:mm:ss",
            '_($* #,##0_);_($* (#,##0);_($* "-"_);_(@_)', '_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)',
            '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)', '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)',
            "[$-x-sysdate]dddd, mmmm dd, yyyy", "[$-x-systime]h:mm:ss AM/PM", "yyyy-mm-dd", "m/d;@", "m/d/yy;@",
            "mm/dd/yy;@", "[$-409]mmmm d, yyyy;@", "dddd, mmmm d, yyyy"]
#: Codes Excel may spell another way: escapes, quotes, case, colours, conditions and sections.
SPELLED = ["#,##0 ;(#,##0)", '0.00- "x"', '0.00\\-\\ "x"', '0.00\\_"x"', "\\$#,##0", '"$"#,##0', "0.00\\%",
           "[RED]0", "[red]0;[blue]-0", "MM/DD/YY", "mmm d, yyyy", "dd/mm/yyyy hh:mm:ss", '0.0 "kr"',
           "0.00E+00;-0.00E+00", "[>100]0;[<=100]0.00", "General;General;General", "0;-0;;@", "#,##0,K",
           "0.0%", "#,##0.00 €", "0.00\\ \\(x\\)", "\\T0", "\\10", "0.00\\E", "AM/PM h", "am/pm h", "a/p h",
           "0.0E+0", "?/?", "0 ?/8", '"("0")"', "[$€-2] #,##0.00", "#,##0;(#,##0)", "0.0_)", "0.0*-"]
#: Custom codes, in the order the layouts make them, some also given to a cell first.
CUSTOM = ["0.000", "0.0000", "0.00000", "mmss.0", "#,##0.000", "0.000%"]


def vba(text: str) -> str:
    """A VBA string literal of ``text``."""
    return '"' + text.replace('"', '""') + '"' if text.isascii() else \
        " & ".join(f"ChrW({ord(char)})" if ord(char) > 127 else '"' + char.replace('"', '""') + '"' for char in text)


#: Every printable character, bare and escaped, after a digit, before one, between date codes and after the text
#: placeholder; then locale and system tags. Each is a rule's format and a cell's on the sheet "chars".
CHARACTER_PROBES = [probe for char in (chr(code) for code in range(33, 127)) if char not in '"'
                    for probe in ("0" + char, char + "0", "d" + char + "m", "@" + char,
                                  "0\\" + char, "d\\" + char + "m", "@\\" + char)]
CHARACTER_PROBES += ["[$-x-sysdate]dddd", "[$-x-systime]h:mm", "[$-F800]dddd", "[$-F400]h:mm", "[$-409]d",
                     "[$-en-US]d", "[$-1009]d", "[$€-x-euro2] #,##0", "0 é", "0é", "0£",
                     "dém", "@é"]


#: name -> (setup, the code the rule's NumberFormat is given); each rule over B2:B3 of the layout's own sheet.
LAYOUTS: dict[str, tuple[str, str]] = {
    **{f"b_{index}": ("", code) for index, code in enumerate(BUILT_IN)},
    **{f"s_{index}": ("", code) for index, code in enumerate(SPELLED)},
    # A custom code a cell holds before a rule takes it, and one only rules hold.
    "c_cell": ('ws.Range("A1").NumberFormat = "0.0000"\n', "0.0000"),
    **{f"c_{index}": ("", code) for index, code in enumerate(CUSTOM)},
    "c_cell_after": ('ws.Range("A1").NumberFormat = "0.000000"\n', "0.00"),
    "c_again": ("", "0.000"),
}


def case_code(index: int, setup: str, code: str) -> str:
    lines = [f"Private Function Case{index}(ws As Object) As String",
             "Dim f As Object, out As String",
             "On Error Resume Next", "Err.Clear", *setup.splitlines(),
             'Set f = ws.Range("B2:B3").FormatConditions.Add(xlCellValue, xlGreater, "0")',
             "f.Interior.Color = RGB(255, 0, 0)", "Err.Clear", f"f.NumberFormat = {vba(code)}",
             'If Err.Number <> 0 Then out = "E" & Err.Number & ";"', "Err.Clear",
             'out = out & TypeName(f.NumberFormat) & ":" & f.NumberFormat',
             'If Err.Number <> 0 Then out = out & "R" & Err.Number', "On Error GoTo 0",
             f"Case{index} = out", "End Function"]
    return "\n".join(lines) + "\n"


#: Character probes to a workbook.
CHUNK = 120


def character_code(start: int, probes: list[str], *, name: str) -> str:
    """A batch of character probes from row ``start`` + 1, each a rule's format on column B and a cell's on column
    A of its row."""
    lines = [f"Private Function {name}(ws As Object) As String",
             "Dim f As Object, out As String, a As String, b As String", "On Error Resume Next"]
    for row, code in enumerate(probes, start + 1):
        lines += ["Err.Clear", f'Set f = ws.Cells({row}, 2).FormatConditions.Add(xlCellValue, xlGreater, "0")',
                  f"f.NumberFormat = {vba(code)}",
                  'If Err.Number <> 0 Then a = "E" & Err.Number Else a = f.NumberFormat', "Err.Clear",
                  f"ws.Cells({row}, 1).NumberFormat = {vba(code)}",
                  f'If Err.Number <> 0 Then b = "E" & Err.Number Else b = ws.Cells({row}, 1).NumberFormat',
                  'out = out & a & "<=>" & b & "<<>>"']
    lines += ["On Error GoTo 0", f"{name} = out", "End Function"]
    return "\n".join(lines) + "\n"


def character_formats(package: zipfile.ZipFile, count: int) -> list[tuple[str, str]]:
    """Each character probe's row: its rule's numFmt and its cell's, as the file keeps them."""
    workbook = package.read("xl/workbook.xml").decode("utf-8")
    targets = {found.group(1): found.group(2) for found in re.finditer(
        r'<Relationship\b[^>]*?Id="([^"]+)"[^>]*?Target="([^"]+)"', package.read("xl/_rels/workbook.xml.rels").decode())}
    rid = re.search(r'<sheet\b[^>]*?name="chars"[^>]*?r:id="([^"]+)"', workbook)
    assert rid is not None
    target = targets[rid.group(1)].lstrip("/")
    text = package.read(target if target.startswith("xl/") else "xl/" + target).decode("utf-8")
    styles = package.read("xl/styles.xml").decode("utf-8")
    dxfs = re.findall(r"<dxf>.*?</dxf>", re.search(r"<dxfs\b.*?</dxfs>", styles, re.DOTALL).group())
    listed = re.search(r"<numFmts\b.*?</numFmts>", styles, re.DOTALL)
    codes = {found.group(1): found.group(0) for found in
             re.finditer(r'<numFmt numFmtId="(\d+)"[^>]*/>', listed.group() if listed else "")}
    xfs = re.findall(r"<xf\b[^>]*?>", re.search(r"<cellXfs\b.*?</cellXfs>", styles, re.DOTALL).group())
    rules = {int(found.group(1)): int(found.group(2)) for found in re.finditer(
        r'<conditionalFormatting sqref="B(\d+)"><cfRule\b[^>]*?dxfId="(\d+)"', text)}
    cells = {int(found.group(1)): int(found.group(2)) for found in re.finditer(r'<c r="A(\d+)" s="(\d+)"', text)}
    out: list[tuple[str, str]] = []
    for row in range(1, count + 1):
        dxf = dxfs[rules[row]] if row in rules else ""
        rule = re.search(r"<numFmt [^>]*/>", dxf)
        xf = xfs[cells[row]] if row in cells else ""
        identifier = re.search(r'numFmtId="(\d+)"', xf)
        number = identifier.group(1) if identifier else ""
        out.append((rule.group() if rule else "", codes.get(number, f'<numFmt numFmtId="{number}"/>')))
    return out


READER = '''Private Function Shown(ws As Object) As String
    Dim f As Object
    On Error Resume Next
    Set f = ws.Range("B2").FormatConditions(1)
    Shown = TypeName(f.NumberFormat) & ":" & f.NumberFormat
    If Err.Number <> 0 Then Shown = "R" & Err.Number
End Function
'''


def number_formats(package: zipfile.ZipFile) -> tuple[dict[str, str], list[str]]:
    """Each sheet's first rule's numFmt as the file keeps it, and the stylesheet's own numFmts."""
    workbook = package.read("xl/workbook.xml").decode("utf-8")
    targets = {found.group(1): found.group(2) for found in re.finditer(
        r'<Relationship\b[^>]*?Id="([^"]+)"[^>]*?Target="([^"]+)"', package.read("xl/_rels/workbook.xml.rels").decode())}
    styles = package.read("xl/styles.xml").decode("utf-8")
    listed = re.search(r"<dxfs\b.*?</dxfs>", styles, re.DOTALL)
    dxfs = re.findall(r"<dxf>.*?</dxf>", listed.group()) if listed else []
    cells = re.findall(r"<numFmt [^>]*/>", (re.search(r"<numFmts\b.*?</numFmts>", styles, re.DOTALL) or [""])[0]
                       if re.search(r"<numFmts\b", styles) else "")
    found: dict[str, str] = {}
    for name, rid in re.findall(r'<sheet\b[^>]*?name="([^"]+)"[^>]*?r:id="([^"]+)"', workbook):
        target = targets[rid].lstrip("/")
        text = package.read(target if target.startswith("xl/") else "xl/" + target).decode("utf-8")
        rule = re.search(r'<cfRule\b[^>]*?dxfId="(\d+)"', text)
        dxf = dxfs[int(rule.group(1))] if rule else ""
        number = re.search(r"<numFmt [^>]*/>", dxf)
        found[name] = number.group() if number else ""
    return found, cells


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    path = OUT / "conditional_number_formats.xlsx"
    again = OUT / "conditional_number_formats_resaved.xlsx"
    for one in (path, again):
        if one.exists():
            one.unlink()
    build = [READER, "Public Function Build() As String", "Dim wb As Object, ws As Object, out As String, f As Object",
             "Application.DisplayAlerts = False", "Set wb = Workbooks.Add(xlWBATWorksheet)"]
    bodies: list[str] = []
    for index, (name, (setup, code)) in enumerate(LAYOUTS.items()):
        build += ["Set ws = wb.Worksheets(1)" if index == 0 else
                  "Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))",
                  f'ws.Name = "{name}"', f'out = out & Case{index}(ws) & "|"']
        bodies.append(case_code(index, setup, code))
    build += [f'wb.SaveAs "{path}", 51', "wb.Close False", 'out = out & "<<PHASE>>"',
              f'Set wb = Workbooks.Open("{path}")', "For Each ws In wb.Worksheets",
              '    out = out & ws.Name & "^" & Shown(ws) & "|"', "Next",
              # A rule made once the file is open takes a custom format no rule had.
              "Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))", 'ws.Name = "after_open"',
              'Set f = ws.Range("B2:B3").FormatConditions.Add(xlCellValue, xlGreater, "0")',
              'f.NumberFormat = "0.0000000"', 'out = out & "after_open^" & Shown(ws) & "|"',
              f'wb.SaveAs "{again}", 51', "wb.Close False"]
    # The character probes go in workbooks of their own, as a workbook holds about 250 custom formats at most.
    for chunk, start in enumerate(range(0, len(CHARACTER_PROBES), CHUNK)):
        build += ['out = out & "<<PHASE>>"', "Set wb = Workbooks.Add(xlWBATWorksheet)", "Set ws = wb.Worksheets(1)",
                  'ws.Name = "chars"']
        for first in range(start, min(start + CHUNK, len(CHARACTER_PROBES)), 100):
            build += [f"out = out & Chars{first}(ws)"]
            bodies.append(character_code(first - start, CHARACTER_PROBES[first:min(first + 100, start + CHUNK)],
                                         name=f"Chars{first}"))
        build += [f'wb.SaveAs "{OUT / f"chars_{chunk}.xlsx"}", 51', "wb.Close False"]
    build += ["Build = out", "End Function"]
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba("\n".join(build) + "\n" + "".join(bodies), "Build", timeout=900.0)
        assert result.ok, f"{result.outcome}: {result.message} {result.error}"
    strip_save_path(path)
    strip_save_path(again)
    before, after, *chunks = str(result.value).split("<<PHASE>>")
    # An answer holding the separator would shift every answer after it onto the wrong layout.
    assert before.count("|") == len(LAYOUTS), "an answer holds the layout separator"
    answers = before.split("|")[: len(LAYOUTS)]
    shown = [one.split("<=>") for chunk in chunks for one in chunk.split("<<>>")[:-1]]
    assert len(shown) == len(CHARACTER_PROBES) and all(len(one) == 2 for one in shown)
    reopened = dict(one.split("^", 1) for one in after.split("|") if one and "^" in one)
    with zipfile.ZipFile(path) as package:
        first, first_cells = number_formats(package)
    with zipfile.ZipFile(again) as package:
        second, second_cells = number_formats(package)
    spelled: list[tuple[str, str]] = []
    for chunk, start in enumerate(range(0, len(CHARACTER_PROBES), CHUNK)):
        strip_save_path(OUT / f"chars_{chunk}.xlsx")
        with zipfile.ZipFile(OUT / f"chars_{chunk}.xlsx") as package:
            spelled += character_formats(package, min(CHUNK, len(CHARACTER_PROBES) - start))
    record = {"reader": READER, "cells": {"file": first_cells, "resaved": second_cells},
              "characters": [{"code": code, "rule": rule, "cell": cell, "rule_file": rule_file, "cell_file": cell_file}
                             for code, (rule, cell), (rule_file, cell_file)
                             in zip(CHARACTER_PROBES, shown, spelled, strict=True)],
              "layouts": [{"name": name, "setup": setup, "code": code, "answer": answer,
                           "reopened": reopened.get(name, ""), "file": first.get(name, ""),
                           "resaved": second.get(name, "")}
                          for (name, (setup, code)), answer in zip(LAYOUTS.items(), answers, strict=True)],
              "after_open": {"answer": reopened.get("after_open", ""), "resaved": second.get("after_open", "")}}
    (OUT / "conditional_number_formats.json").write_text(json.dumps(record, indent=1, ensure_ascii=False) + "\n",
                                                         encoding="utf-8")
    for one in record["layouts"]:
        print(f"{one['name']:6} {one['code']!r:48} {one['answer']!r:44} {one['file']} -> {one['resaved']}")
        if one["reopened"] != one["answer"]:
            print(f"       reopened {one['reopened']!r}")
    print("cells", first_cells, "->", second_cells)
    print("after open", record["after_open"])


if __name__ == "__main__":
    main()
