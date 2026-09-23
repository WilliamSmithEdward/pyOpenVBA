"""How tall Excel makes a row for the fonts in its cells, on a 96-DPI display.

A row that keeps no height of its own is as tall as its tallest font, and
never shorter than the standard row. How tall a font makes it depends on
the font's pixel size, round(points * 4 / 3), and on the font itself,
through its hinted metrics, irregularly enough that no formula over the
font's tables reproduces it. So this measures every pixel size from 1 to
546 (409.5pt) for each font, in regular, bold, italic and bold italic,
and records the row's height in pixels and the descent Excel writes for
it. Excel keeps about a thousand fonts in a workbook, so each batch of
cases gets a workbook of its own.

    python scripts/measure_font_heights.py

writes tests/fixtures/font_rows.json, from which
scripts/bake_font_rows.py builds src/pyopenvba/apps/excel/_font_rows.py.
"""

from __future__ import annotations

import json
import re
import tempfile
import zipfile
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "tests" / "fixtures" / "font_rows.json"
#: The fonts measured. Aptos Display and Aptos Narrow share Aptos's heights
#: and Calibri Light Calibri's, which the regular measurements confirm.
FONTS = ("Aptos", "Aptos Display", "Aptos Narrow", "Arial", "Calibri", "Calibri Light", "Cambria", "Consolas",
         "Courier New", "Georgia", "Segoe UI", "Tahoma", "Times New Roman", "Verdana")
SHARING = {"Aptos Display", "Aptos Narrow", "Calibri Light"}
VARIANTS = ((False, False), (True, False), (False, True), (True, True))
PIXELS = range(1, 547)
BATCH = 400


def cases() -> list[tuple[str, bool, bool, int]]:
    out: list[tuple[str, bool, bool, int]] = []
    for font in FONTS:
        for bold, italic in VARIANTS[:1] if font in SHARING else VARIANTS:
            out += [(font, bold, italic, pixels) for pixels in PIXELS]
    return out


def points(pixels: int) -> float:
    """A point size that is exactly this many pixels; Excel refuses sizes below 1pt, which is 1 pixel too."""
    return max(pixels * 0.75, 1.0)


def procedure(index: int, batch: list[tuple[str, bool, bool, int]], target: Path) -> list[str]:
    spec = ";".join(f"{font}|{int(bold)}|{int(italic)}|{points(pixels)!r}" for font, bold, italic, pixels in batch)
    pieces = [spec[start:start + 800] for start in range(0, len(spec), 800)]
    lines = [f"Private Function Batch{index}() As String", "Dim out As String, i As Long, spec As Variant",
             "Dim parts As Variant, s As String, wb As Object, ws As Object",
             "Set wb = Workbooks.Add(xlWBATWorksheet)", "Set ws = wb.Worksheets(1)"]
    lines += [f's = s & "{piece}"' for piece in pieces]
    lines += ['spec = Split(s, ";")', "For i = 0 To UBound(spec)", 'parts = Split(spec(i), "|")',
              "With ws.Cells(i + 1, 1).Font", ".Name = parts(0)", '.Bold = (parts(1) = "1")',
              '.Italic = (parts(2) = "1")', ".Size = Val(parts(3))", "End With",
              'out = out & ws.Rows(i + 1).Height & ","', "Next i",
              f'wb.SaveAs Filename:="{target}", FileFormat:=51', "wb.Close False",
              f"Batch{index} = out", "End Function"]
    return lines


#: Workbooks a single run builds, so the probe module stays well inside VBA's limits.
PER_RUN = 8


def main() -> None:
    items = cases()
    folder = Path(tempfile.mkdtemp())
    batches = [items[start:start + BATCH] for start in range(0, len(items), BATCH)]
    targets = [folder / f"rows{index}.xlsx" for index in range(len(batches))]
    blocks: list[str] = []
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        standard = excel.run_vba("Public Function S() As String\nS = Workbooks.Add(xlWBATWorksheet).Worksheets(1)"
                                 ".StandardHeight\nActiveWorkbook.Close False\nEnd Function\n", "S", timeout=60.0)
        assert standard.ok and str(standard.value) == "15", f"needs a 96-DPI display ({standard.value})"
        for first in range(0, len(batches), PER_RUN):
            indexes = range(first, min(first + PER_RUN, len(batches)))
            head = ["Public Function Probe() As String", "Dim out As String", "Application.DisplayAlerts = False",
                    "Application.ScreenUpdating = False"]
            body: list[str] = []
            for index in indexes:
                head.append(f'out = out & Batch{index}() & "#"')
                body += procedure(index, batches[index], targets[index])
            head += ["Application.ScreenUpdating = True", "Probe = out", "End Function"]
            result = excel.run_vba("\n".join(head + body) + "\n", "Probe", timeout=900.0)
            assert result.ok, f"{result.outcome}: {result.message} {result.error}"
            blocks += str(result.value).split("#")[: len(indexes)]
            print(f"batches {first}-{indexes[-1]} of {len(batches)}", flush=True)
    records: list[dict[str, object]] = []
    for batch, block, target in zip(batches, blocks, targets, strict=True):
        with zipfile.ZipFile(target) as package:
            xml = package.read("xl/worksheets/sheet1.xml").decode("utf-8")
        heights = [item for item in block.split(",") if item]
        tags = {int(m.group(1)): m.group(0) for m in re.finditer(r'<row r="(\d+)"[^>]*>', xml)}
        for row, ((font, bold, italic, pixels), height) in enumerate(zip(batch, heights, strict=True), start=1):
            tag = tags.get(row, "")
            descent = re.search(r'dyDescent="([^"]*)"', tag)
            stored = re.search(r'\bht="([^"]*)"', tag)
            records.append({"font": font, "bold": bold, "italic": italic, "pixels": pixels,
                            "row": round(float(height) / 0.75), "ht": stored.group(1) if stored else None,
                            "descent": descent.group(1) if descent else None})
    tables = compact(records)
    OUT.write_text(written(tables), encoding="utf-8")
    print("saved", len(tables), "tables from", len(records), "cases")


def written(tables: list[dict[str, object]]) -> str:
    """The fixture as JSON, one table to a line."""
    return "[\n" + ",\n".join(json.dumps(table, separators=(",", ":")) for table in tables) + "\n]\n"


def compact(records: list[dict[str, object]]) -> list[dict[str, object]]:
    """One table per font and style: the row's pixels and the descent in twips, for pixel sizes 1 to 546.

    What Excel wrote is kept only as far as it says more than those
    numbers: every ht it wrote is the height rule's spelling of the row,
    every descent the measurement spelling of its twips, and a row as
    tall as the standard one has no ht. This checks all three.
    """
    import sys

    sys.path.insert(0, str(ROOT / "src"))
    from pyopenvba.apps.excel._dimensions import height_text
    from pyopenvba.apps.excel._styles import excel_number

    grouped: dict[tuple[str, bool, bool], dict[int, dict[str, object]]] = {}
    for record in records:
        key = (str(record["font"]), bool(record["bold"]), bool(record["italic"]))
        grouped.setdefault(key, {})[int(record["pixels"])] = record  # type: ignore[call-overload]
    out: list[dict[str, object]] = []
    for (font, bold, italic), by_pixels in grouped.items():
        rows: list[int] = []
        descents: list[int] = []
        for pixels in PIXELS:
            record = by_pixels[pixels]
            row = int(record["row"])  # type: ignore[call-overload]
            spelled = str(record["descent"])
            twips = round(float(spelled) * 20)
            assert excel_number(twips / 20) == spelled, (font, pixels, spelled)
            assert record["ht"] == (height_text(row * 4) if row > 20 else None), (font, pixels, record["ht"])
            rows.append(row)
            descents.append(twips)
        out.append({"font": font, "bold": bold, "italic": italic, "rows": rows, "descent": descents})
    return out


if __name__ == "__main__":
    main()
