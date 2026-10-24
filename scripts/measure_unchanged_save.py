"""What Excel writes of a stylesheet when it opens a file and saves it with nothing changed.

Excel writes what is in use and nothing else whenever it saves. This
makes a workbook in Excel with cells in a format each, a row's and a
column's format, a cell in the built-in style Good, a style of a user's
that nothing uses and a conditional format's dxf, then writes copies of
it holding what Excel would not write: an xf nothing uses, in the middle
of the table and at its end; a font, a fill and a border nothing uses;
an unused xf with parts only it names; an xf like an earlier one, with
the same flags and with others; a font twice; a built-in style nothing
uses; an xf only a row uses; and a dxf no rule names. Excel opens each,
reads its cells' formats, and saves it again with nothing changed.

    python scripts/measure_unchanged_save.py

writes tests/fixtures/unchanged_save/: the workbook, each copy, Excel's
save of each, and unchanged_save.json, which
tests/test_excel_unchanged_save.py replays.
"""

from __future__ import annotations

import json
import re
import shutil
import sys
import zipfile
from collections.abc import Callable
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "scripts"))
from fixture_workbook import strip_save_path  # noqa: E402

OUT = ROOT / "tests" / "fixtures" / "unchanged_save"
#: The stylesheet's sections, as a save writes them.
SECTIONS = ("numFmts", "fonts", "fills", "borders", "cellStyleXfs", "cellXfs", "cellStyles", "dxfs")
_SECTION = {"numFmts": "numFmt", "fonts": "font", "fills": "fill", "borders": "border", "cellStyleXfs": "xf",
            "cellXfs": "xf", "cellStyles": "cellStyle", "dxfs": "dxf"}

BUILD = '''Public Function Build() As String
Dim wb As Object, ws As Object, f As Object
Application.DisplayAlerts = False
Set wb = Workbooks.Add(xlWBATWorksheet)
Set ws = wb.Worksheets(1)
ws.Name = "cells"
ws.Range("A1:A6").Value = 1
ws.Range("A1").Font.Bold = True
ws.Range("A2").Interior.Color = RGB(255, 0, 0)
ws.Range("A3").Borders(xlEdgeBottom).LineStyle = xlContinuous
ws.Range("A4").NumberFormat = "0.000"
ws.Range("A5").Font.Italic = True
ws.Range("A5").Interior.Color = RGB(0, 0, 255)
ws.Range("A6").Style = "Good"
ws.Rows(8).Font.Underline = xlUnderlineStyleSingle
ws.Columns("D").Interior.Color = RGB(0, 255, 0)
wb.Styles.Add "Mine"
wb.Styles("Mine").Font.Size = 14
ws.Range("B2:B3").Value = 5
Set f = ws.Range("B2:B3").FormatConditions.Add(xlCellValue, xlGreater, "0")
f.Interior.Color = RGB(255, 255, 0)
wb.SaveAs "{path}", 51
wb.Close False
Build = "ok"
End Function
'''

READ = '''Private Function Look(c As Object) As String
    On Error Resume Next
    Look = c.Address(False, False) & "=" & c.Font.Bold & "/" & c.Font.Italic & "/" & c.Font.Underline & "/" & _
        c.Font.Size & "/" & c.Interior.Color & "/" & c.Borders(xlEdgeBottom).LineStyle & "/" & c.NumberFormat & _
        "/" & c.Style.Name
    If Err.Number <> 0 Then Look = c.Address(False, False) & "=E" & Err.Number
End Function

Public Function Read() As String
Dim wb As Object, ws As Object, out As String, name As Variant, c As Object, s As Object
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
        For Each c In ws.Range("A1:A6,B1,A8,D1,A10")
            out = out & Look(c) & ";"
        Next
        For Each s In wb.Styles
            If Not s.BuiltIn Then out = out & "style:" & s.Name & ";"
        Next
        out = out & "styles:" & wb.Styles.Count & "|"
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


def entries(styles: str, section: str) -> list[str]:
    """A stylesheet section's entries as the file writes them, without the random uids Excel gives styles."""
    found = re.search(rf"<{section}\b[^>]*?(?:/>|>(.*?)</{section}>)", styles, re.DOTALL)
    if found is None or found.group(1) is None:
        return []
    tag = _SECTION[section]
    listed = re.findall(rf"<{tag}\b[^>]*?/>|<{tag}\b[^>]*?>.*?</{tag}>", found.group(1), re.DOTALL)
    return [re.sub(r' xr:uid="[^"]*"', "", one) for one in listed]


def holdings(path: Path) -> dict[str, object]:
    """What a file holds of its formats: each stylesheet section's entries, and the xf each of the sheet "cells"'
    cells, rows and columns points at."""
    with zipfile.ZipFile(path) as package:
        styles = package.read("xl/styles.xml").decode("utf-8")
        sheet = package.read(_part(package, "cells")).decode("utf-8")
    return {**{section: entries(styles, section) for section in SECTIONS},
            "cells": dict(re.findall(r'<c r="([A-Z]+\d+)" s="(\d+)"', sheet)),
            "rows": dict(re.findall(r'<row r="(\d+)"[^>]*? s="(\d+)"', sheet)),
            "cols": [f"{low}:{high}:{style}" for low, high, style in
                     re.findall(r'<col min="(\d+)" max="(\d+)"[^>]*? style="(\d+)"', sheet)]}


def _replaced(text: str, section: str, change: Callable[[list[str]], list[str]]) -> str:
    """``text`` with ``section``'s entries as ``change`` makes them, its count following."""
    found = re.search(rf'<{section} count="(\d+)"[^>]*>(.*?)</{section}>', text, re.DOTALL)
    assert found is not None, section
    tag = _SECTION[section]
    listed = re.findall(rf"<{tag}\b[^>]*?/>|<{tag}\b[^>]*?>.*?</{tag}>", found.group(2), re.DOTALL)
    made = change(listed)
    opened = found.group(0)[:found.group(0).index(">") + 1].replace(f'count="{found.group(1)}"', f'count="{len(made)}"')
    return text[:found.start()] + opened + "".join(made) + f"</{section}>" + text[found.end():]


def _shifted(sheet: str, after: int) -> str:
    """``sheet`` with every xf index past ``after`` one up, as an xf put in at ``after`` + 1 moves them."""
    def up(found: re.Match[str]) -> str:
        index = int(found.group(2))
        return found.group(1) + str(index + 1 if index > after else index) + found.group(3)

    return re.sub(r'(<(?:c|row)\b[^>]*?\bs="|<col\b[^>]*?\bstyle=")(\d+)(")', up, sheet)


def _with_cell(sheet: str, reference: str, xf: int) -> str:
    """``sheet`` with the cell ``reference`` of row 1 holding 1 in the xf ``xf``."""
    def column(name: str) -> int:
        letters = re.match(r"[A-Z]+", name)
        assert letters is not None
        number = 0
        for letter in letters.group():
            number = number * 26 + ord(letter) - 64
        return number

    row = re.search(r'<row r="1"[^>]*>(.*?)</row>', sheet, re.DOTALL)
    assert row is not None
    found = list(re.finditer(r'<c r="([A-Z]+)1"[^>]*?(?:/>|>.*?</c>)', row.group(1), re.DOTALL))
    position = next((one.start() for one in found if column(one.group(1)) > column(reference)), len(row.group(1)))
    cells = row.group(1)[:position] + f'<c r="{reference}" s="{xf}"><v>1</v></c>' + row.group(1)[position:]
    return sheet[:row.start(1)] + cells + sheet[row.end(1):]


def _with_row(sheet: str, number: int, xf: int) -> str:
    """``sheet`` with an empty row ``number`` in the xf ``xf``."""
    rows = list(re.finditer(r'<row r="(\d+)"[^>]*?(?:/>|>.*?</row>)', sheet, re.DOTALL))
    position = next((found.start() for found in rows if int(found.group(1)) > number), rows[-1].end())
    added = f'<row r="{number}" s="{xf}" customFormat="1"/>'
    return sheet[:position] + added + sheet[position:]


def _named(xf: str, attribute: str) -> str:
    """The value an xf gives ``attribute``."""
    found = re.search(rf'{attribute}="(\d+)"', xf)
    assert found is not None, attribute
    return found.group(1)


def variants(styles: str, sheet: str) -> dict[str, tuple[str, str]]:
    """Each copy's stylesheet and sheet."""
    fonts = entries(styles, "fonts")
    fills = entries(styles, "fills")
    borders = entries(styles, "borders")
    xfs = entries(styles, "cellXfs")
    cells = dict(re.findall(r'<c r="([A-Z]+\d+)" s="(\d+)"', sheet))
    bold, italic_blue = xfs[int(cells["A1"])], xfs[int(cells["A5"])]
    italic = _named(italic_blue, "fontId")
    red = _named(xfs[int(cells["A2"])], "fillId")
    stray = f'<xf numFmtId="0" fontId="{italic}" fillId="{red}" borderId="0" xfId="0" applyFont="1" applyFill="1"/>'
    font = '<font><sz val="20"/><color theme="1"/><name val="Aptos Narrow"/><family val="2"/><scheme val="minor"/></font>'
    fill = '<fill><patternFill patternType="solid"><fgColor rgb="FFFFC0CB"/><bgColor indexed="64"/></patternFill></fill>'
    border = '<border><left/><right/><top style="thin"><color indexed="64"/></top><bottom/><diagonal/></border>'
    own = (f'<xf numFmtId="0" fontId="{len(fonts)}" fillId="{len(fills)}" borderId="{len(borders)}" xfId="0" '
           'applyFont="1" applyFill="1" applyBorder="1"/>')
    flagless = re.sub(r' applyFont="1"', "", bold)
    second_bold = re.sub(r'fontId="\d+"', f'fontId="{len(fonts)}"', bold)
    bold_font = fonts[int(_named(bold, "fontId"))]
    neutral_font = '<font><sz val="11"/><color rgb="FF9C5700"/><name val="Aptos Narrow"/><family val="2"/>' \
                   '<scheme val="minor"/></font>'
    neutral_fill = '<fill><patternFill patternType="solid"><fgColor rgb="FFFFEB9C"/></patternFill></fill>'
    neutral = (f'<xf numFmtId="0" fontId="{len(fonts)}" fillId="{len(fills)}" borderId="0" applyNumberFormat="0" '
               'applyBorder="0" applyAlignment="0" applyProtection="0"/>')
    style_count = len(entries(styles, "cellStyleXfs"))
    unused_dxf = '<dxf><font><color rgb="FF9C0006"/></font></dxf>'
    count = len(xfs)
    return {
        "base": (styles, sheet),
        "unused_middle": (_replaced(styles, "cellXfs", lambda listed: [*listed[:2], stray, *listed[2:]]),
                          _shifted(sheet, 1)),
        "unused_end": (_replaced(styles, "cellXfs", lambda listed: [*listed, stray]), sheet),
        "unused_parts": (_replaced(_replaced(_replaced(styles, "fonts", lambda listed: [*listed, font]), "fills",
                                             lambda listed: [*listed, fill]), "borders",
                                   lambda listed: [*listed, border]), sheet),
        "unused_xf_parts": (_replaced(_replaced(_replaced(_replaced(styles, "fonts", lambda listed: [*listed, font]),
                                                          "fills", lambda listed: [*listed, fill]), "borders",
                                                lambda listed: [*listed, border]), "cellXfs",
                                      lambda listed: [*listed, own]), sheet),
        "alike": (_replaced(styles, "cellXfs", lambda listed: [*listed, bold]), _with_cell(sheet, "B1", count)),
        "alike_flags": (_replaced(styles, "cellXfs", lambda listed: [*listed, flagless]),
                        _with_cell(sheet, "B1", count)),
        "duplicate_font": (_replaced(_replaced(styles, "fonts", lambda listed: [*listed, bold_font]), "cellXfs",
                                     lambda listed: [*listed, second_bold]), _with_cell(sheet, "B1", count)),
        "unused_builtin_style": (_replaced(_replaced(_replaced(_replaced(
            styles, "fonts", lambda listed: [*listed, neutral_font]), "fills", lambda listed: [*listed, neutral_fill]),
            "cellStyleXfs", lambda listed: [*listed, neutral]), "cellStyles",
            lambda listed: [*listed, f'<cellStyle name="Neutral" xfId="{style_count}" builtinId="28"/>']), sheet),
        "row_only": (_replaced(styles, "cellXfs", lambda listed: [*listed, stray]), _with_row(sheet, 10, count)),
        "unused_dxf": (_replaced(styles, "dxfs", lambda listed: [*listed, unused_dxf]), sheet),
    }


def write(base: Path, target: Path, styles: str, sheet: str, sheet_part: str) -> None:
    """``base`` with the stylesheet ``styles`` and the sheet "cells" ``sheet``."""
    with zipfile.ZipFile(base) as package, zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as out:
        for info in package.infolist():
            data = package.read(info.filename)
            if info.filename == "xl/styles.xml":
                data = styles.encode("utf-8")
            elif info.filename == sheet_part:
                data = sheet.encode("utf-8")
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
        sheet_part = _part(package, "cells")
        sheet = package.read(sheet_part).decode("utf-8")
    made = variants(styles, sheet)
    for name, (styles_text, sheet_text) in made.items():
        if name != "base":
            write(base, OUT / f"{name}.xlsx", styles_text, sheet_text, sheet_part)
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
        cases.append({"name": name, "answer": answers.get(name, ""), "file": holdings(OUT / f"{name}.xlsx"),
                      "resaved": holdings(resaved) if resaved.exists() else {}})
    (OUT / "unchanged_save.json").write_text(json.dumps({"read": READ, "cases": cases}, indent=1,
                                                        ensure_ascii=False) + "\n", encoding="utf-8")
    for one in cases:
        print("==", one["name"], ascii(one["answer"]))
        file, again = one["file"], one["resaved"]
        assert isinstance(file, dict) and isinstance(again, dict)
        for key in file:
            if file[key] != again.get(key):
                print(f"   {key}")
                print(f"      file    {ascii(file[key])}")
                print(f"      resaved {ascii(again.get(key))}")


if __name__ == "__main__":
    main()
