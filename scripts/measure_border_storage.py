"""Where Excel stores a border it is asked to set, and how neighbouring cells read it back.

A border between two cells can be stored on either of them. Each case
here sets borders in its own part of a fresh sheet, saves the workbook as
.xlsx, and records two things: the edges Excel reads for the cells around
the change, and the sides each of those cells stores in the saved file.
The fill patterns ride along: every XlPattern set on one cell, with the
patternType the file stores for it.

    python scripts/measure_border_storage.py

writes tests/fixtures/format/border_storage.json.
"""

from __future__ import annotations

import json
import re
import tempfile
import zipfile
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parent.parent

#: name -> (setup, cells whose edges are read and whose stored sides are recorded)
CASES: dict[str, tuple[str, str]] = {
    "bottom": ('Range("B2").Borders(xlEdgeBottom).LineStyle = xlContinuous', "B2 B3"),
    "top": ('Range("D2").Borders(xlEdgeTop).LineStyle = xlContinuous', "D1 D2"),
    "right": ('Range("F2").Borders(xlEdgeRight).LineStyle = xlContinuous', "F2 G2"),
    "left": ('Range("J2").Borders(xlEdgeLeft).LineStyle = xlContinuous', "I2 J2"),
    "range_top": ('Range("B5:C6").Borders(xlEdgeTop).LineStyle = xlContinuous', "B4 C4 B5 C5 B6"),
    "range_bottom": ('Range("E5:F6").Borders(xlEdgeBottom).LineStyle = xlContinuous', "E6 F6 E7 F7"),
    "inside_horizontal": ('Range("B9:C10").Borders(xlInsideHorizontal).LineStyle = xlContinuous', "B9 C9 B10 C10"),
    "inside_vertical": ('Range("E9:F10").Borders(xlInsideVertical).LineStyle = xlContinuous', "E9 F9 E10 F10"),
    "all": ('Range("H9:I10").Borders.LineStyle = xlContinuous', "H9 I9 H10 I10 H11 J9"),
    "around": ('Range("B13:C14").BorderAround xlContinuous, xlMedium', "B12 B13 C13 B14 C14 B15 D13 A13"),
    "conflict_bottom_then_top": ('Range("B17").Borders(xlEdgeBottom).LineStyle = xlContinuous\n'
                                 'Range("B18").Borders(xlEdgeTop).Weight = xlThick', "B17 B18"),
    "conflict_top_then_bottom": ('Range("D18").Borders(xlEdgeTop).Weight = xlThick\n'
                                 'Range("D17").Borders(xlEdgeBottom).LineStyle = xlDash', "D17 D18"),
    "conflict_right_then_left": ('Range("F17").Borders(xlEdgeRight).LineStyle = xlContinuous\n'
                                 'Range("G17").Borders(xlEdgeLeft).LineStyle = xlDouble', "F17 G17"),
    "clear_from_below": ('Range("B21").Borders(xlEdgeBottom).LineStyle = xlContinuous\n'
                         'Range("B22").Borders(xlEdgeTop).LineStyle = xlNone', "B21 B22"),
    "clear_from_above": ('Range("D22").Borders(xlEdgeTop).LineStyle = xlContinuous\n'
                         'Range("D21").Borders(xlEdgeBottom).LineStyle = xlNone', "D21 D22"),
    "clear_own": ('Range("F21").Borders(xlEdgeBottom).LineStyle = xlContinuous\n'
                  'Range("F21").Borders(xlEdgeBottom).LineStyle = xlNone', "F21 F22"),
    "neighbour_font_change": ('Range("B25").Borders(xlEdgeBottom).LineStyle = xlContinuous\n'
                              'Range("B26").Font.Bold = True', "B25 B26"),
    "neighbour_border_change": ('Range("D25").Borders(xlEdgeBottom).LineStyle = xlContinuous\n'
                                'Range("D26").Borders(xlEdgeLeft).LineStyle = xlContinuous', "D25 D26"),
    "neighbour_left_change": ('Range("F25").Borders(xlEdgeRight).LineStyle = xlContinuous\n'
                              'Range("G25").Borders(xlEdgeTop).LineStyle = xlContinuous', "F25 G25"),
    "neighbour_above_change": ('Range("I26").Borders(xlEdgeTop).LineStyle = xlContinuous\n'
                               'Range("I25").Borders(xlEdgeLeft).LineStyle = xlContinuous', "I25 I26"),
    "clear_formats_neighbour": ('Range("B29").Borders(xlEdgeBottom).LineStyle = xlContinuous\n'
                                'Range("B29").ClearFormats', "B29 B30"),
    "clear_formats_owner_below": ('Range("D30").Borders(xlEdgeTop).LineStyle = xlContinuous\n'
                                  'Range("D29").ClearFormats', "D29 D30"),
    "range_inside_existing": ('Range("B33").Borders(xlEdgeBottom).LineStyle = xlDouble\n'
                              'Range("B33:B34").Borders(xlInsideHorizontal).LineStyle = xlContinuous', "B33 B34"),
    "diagonal_neighbours": ('Range("D33").Borders(xlDiagonalDown).LineStyle = xlContinuous', "D33 E33 D34"),
}

PATTERNS = (-4142, 1, -4126, -4125, -4124, 17, 18, -4128, -4166, -4121, -4162, 9, 10, 11, 12, 13, 14, 15, 16)


def main() -> None:
    folder = Path(tempfile.mkdtemp())
    target = folder / "border_storage.xlsx"
    setup = ["Public Function Build() As String", "Dim wb As Object, ws As Object, c As Object, e As Variant",
             "Application.DisplayAlerts = False", "Set wb = Workbooks.Add(xlWBATWorksheet)",
             "Set ws = wb.Worksheets(1)", "ws.Activate"]
    for statements, _ in CASES.values():
        setup += [("ws." + line) if line.startswith("Range") else line for line in statements.splitlines()]
    for row, pattern in enumerate(PATTERNS, start=40):
        setup.append(f'ws.Range("B{row}").Interior.Pattern = {pattern}')
    # Read every edge of every recorded cell, then save.
    setup.append('Build = ""')
    for name, (_, cells) in CASES.items():
        for cell in cells.split():
            setup.append(f'Build = Build & "{name}.{cell}="')
            setup.append("For Each e In Array(5, 6, 7, 8, 9, 10)")
            setup.append(f'Build = Build & ws.Range("{cell}").Borders(e).LineStyle & ":" & '
                         f'ws.Range("{cell}").Borders(e).Weight & ","')
            setup.append("Next e")
            setup.append('Build = Build & "|"')
    for row in range(40, 40 + len(PATTERNS)):
        setup.append(f'Build = Build & "pattern.B{row}=" & ws.Range("B{row}").Interior.Pattern & "|"')
    setup += [f'wb.SaveAs Filename:="{target}", FileFormat:=51', "wb.Close False", "End Function"]
    with ExcelSession(HarnessConfig(lock_wait_s=45.0)) as excel:
        excel.new_document()
        result = excel.run_vba("\n".join(setup) + "\n", "Build", timeout=240.0)
        assert result.ok, f"{result.outcome}: {result.message}"
    reads = dict(item.split("=", 1) for item in str(result.value).split("|") if item)
    with zipfile.ZipFile(target) as package:
        styles = package.read("xl/styles.xml").decode("utf-8")
        sheet = package.read("xl/worksheets/sheet1.xml").decode("utf-8")
    xfs = re.findall(r"<xf\b[^>]*/>|<xf\b[^>]*>.*?</xf>",
                     re.search(r"<cellXfs.*?</cellXfs>", styles, re.DOTALL).group(0))
    borders = re.findall(r"<border\b[^>]*/>|<border\b[^>]*>.*?</border>",
                         re.search(r"<borders.*?</borders>", styles, re.DOTALL).group(0))
    fills = re.findall(r"<fill\b[^>]*>.*?</fill>", re.search(r"<fills.*?</fills>", styles, re.DOTALL).group(0))
    style_of = {ref: int(index) for ref, index in re.findall(r'<c r="([A-Z]+\d+)" s="(\d+)"', sheet)}

    def stored(ref: str) -> str:
        xf = xfs[style_of.get(ref, 0)]
        return borders[int(re.search(r'borderId="(\d+)"', xf).group(1))]

    records = []
    for name, (statements, cells) in CASES.items():
        records.append({
            "name": name,
            "setup": statements,
            "cells": {cell: {"reads": reads[f"{name}.{cell}"], "stored": stored(cell)} for cell in cells.split()},
        })
    patterns = []
    for row, pattern in enumerate(PATTERNS, start=40):
        xf = xfs[style_of.get(f"B{row}", 0)]
        fill = fills[int(re.search(r'fillId="(\d+)"', xf).group(1))]
        patterns.append({"pattern": pattern, "reads": reads[f"pattern.B{row}"], "stored": fill})
    payload = {"cases": records, "patterns": patterns}
    out = ROOT / "tests/fixtures/format/border_storage.json"
    out.write_text(json.dumps(payload, indent=1) + "\n", encoding="utf-8")
    for record in records:
        print(record["name"])
        for cell, found in record["cells"].items():
            sides = re.sub(r"<color[^>]*/>|</?(?:left|right|top|bottom|diagonal)>|<border>|</border>", "", found["stored"])
            print(f"   {cell}: reads {found['reads']}  stored {sides}")
    for item in patterns:
        print(item["pattern"], item["reads"], re.search(r'patternType="([^"]*)"', item["stored"]).group(1))


if __name__ == "__main__":
    main()
