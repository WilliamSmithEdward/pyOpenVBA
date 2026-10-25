"""AutoFit and row heights over rich text, on a 96-DPI display.

A cell whose text holds runs in different fonts (Range.Characters): how
tall Excel makes its row, how wide AutoFit makes its column and how tall
Rows.AutoFit makes its row. Each case sets its runs in A1 of a sheet of
its own, in a workbook whose Normal font is each of the two the model
measures (Aptos Narrow 11 and Calibri 11); the workbook is then saved, to
record the rows and columns the file keeps.

    python scripts/measure_rich_text_fit.py

writes tests/fixtures/rich_text_fit.json, which tests/test_excel_characters.py
replays. Excel must lay sheets out at 96 DPI, a 15pt standard row.
"""
from __future__ import annotations

import json
import re
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / "tests/fixtures/rich_text_fit.json"
#: Each Normal font, and the other measured font a run is set in.
NORMALS = {"Aptos Narrow": "Calibri", "Calibri": "Aptos Narrow"}
#: What each case does to A1, c, which holds abcdefgh; OTHER is the other measured font.
CASES = {
    "bold_first": "c.Characters(1, 3).Font.Bold = True",
    "bold_last": "c.Characters(6, 3).Font.Bold = True",
    "bold_middle": "c.Characters(3, 3).Font.Bold = True",
    "bold_alternate": "For i = 1 To 8 Step 2: c.Characters(i, 1).Font.Bold = True: Next",
    "italic_first": "c.Characters(1, 3).Font.Italic = True",
    "italic_last": "c.Characters(6, 3).Font.Italic = True",
    "italic_middle": "c.Characters(3, 3).Font.Italic = True",
    "italic_after_first": "c.Characters(2, 7).Font.Italic = True",
    "italic_bold_cell": "c.Font.Bold = True: c.Characters(6, 3).Font.Italic = True",
    "bold_italic_mix": "c.Characters(1, 4).Font.Bold = True: c.Characters(3, 4).Font.Italic = True",
    "size8_last": "c.Characters(6, 3).Font.Size = 8",
    "size16_last": "c.Characters(6, 3).Font.Size = 16",
    "size16_first": "c.Characters(1, 3).Font.Size = 16",
    "size30_last": "c.Characters(6, 3).Font.Size = 30",
    "size30_first": "c.Characters(1, 3).Font.Size = 30",
    "size30_middle": "c.Characters(4, 2).Font.Size = 30",
    "size16_italic_last": "c.Characters(6, 3).Font.Size = 16: c.Characters(6, 3).Font.Italic = True",
    "italic_alternate": "For i = 2 To 8 Step 2: c.Characters(i, 1).Font.Italic = True: Next",
    "italic_alternate_first": "For i = 1 To 8 Step 2: c.Characters(i, 1).Font.Italic = True: Next",
    "italic_two_first": "c.Characters(1, 2).Font.Italic = True: c.Characters(6, 2).Font.Italic = True",
    "italic_cell_bold_last": "c.Font.Italic = True: c.Characters(6, 3).Font.Bold = True",
    "size72_italic_last": "c.Font.Size = 72: c.Characters(6, 3).Font.Italic = True",
    "size72_italic_alternate": "c.Font.Size = 72: For i = 2 To 8 Step 2: c.Characters(i, 1).Font.Italic = True: Next",
    "size72_italic_alternate_first": "c.Font.Size = 72: For i = 1 To 8 Step 2: c.Characters(i, 1).Font.Italic = True: Next",
    "size48_italic_middle": "c.Font.Size = 48: c.Characters(3, 3).Font.Italic = True",
    "size72_after_small": "c.Characters(2, 7).Font.Size = 72: For i = 2 To 8 Step 2: c.Characters(i, 1).Font.Italic = True: Next",
    "size16_reset": "c.Characters(6, 3).Font.Size = 16: c.Font.Size = 11",
    "other_font_last": 'c.Characters(6, 3).Font.Name = "OTHER"',
    "other_font_first": 'c.Characters(1, 3).Font.Name = "OTHER"',
    "superscript_last": "c.Characters(6, 3).Font.Superscript = True",
    "subscript_last": "c.Characters(6, 3).Font.Subscript = True",
    "underline_last": "c.Characters(6, 3).Font.Underline = xlUnderlineStyleSingle",
    "strikethrough_last": "c.Characters(6, 3).Font.Strikethrough = True",
    "color_last": "c.Characters(6, 3).Font.Color = RGB(255, 0, 0)",
    "beside_plain": 'c.Characters(6, 3).Font.Size = 16: ws.Range("B1").Value = "x": ws.Range("B1").Font.Size = 20',
    "quoted": "c.Value = \"'abcdef\": c.Characters(1, 3).Font.Bold = True",
}

#: Each case's row height as Excel set it, the column's width AutoFit made in pixels, and the row's height
#: Rows.AutoFit made, joined by #; then the workbook saved to ``path``.
OPENING = '''Function Shape(ws As Worksheet) As String
Dim before As Double
before = ws.Rows(1).Height
ws.Columns(1).AutoFit
Shape = before & "/" & CLng(ws.Columns(1).Width * 4 / 3) & "/"
ws.Rows(1).AutoFit
Shape = Shape & ws.Rows(1).Height
End Function

Function Probe(path As String) As String
Dim wb As Workbook, ws As Worksheet, c As Range, s As String, i As Long
Application.DisplayAlerts = False
Set wb = Workbooks.Add(xlWBATWorksheet)
wb.Styles("Normal").Font.Name = "NORMAL"
wb.Styles("Normal").Font.Size = 11
'''
CASE = '''Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
ws.Name = "NAME"
Set c = ws.Range("A1")
c.Value = "abcdefgh"
SETUP
s = s & "#NAME=" & Shape(ws)
'''
CLOSING = '''wb.SaveAs path, xlOpenXMLWorkbook
wb.Close False
Application.DisplayAlerts = True
Probe = Mid$(s, 2)
End Function
'''


def module(normal: str) -> str:
    """The probe for a workbook whose Normal font is ``normal``."""
    body = "".join(CASE.replace("NAME", name).replace("SETUP", setup.replace("OTHER", NORMALS[normal]))
                   for name, setup in CASES.items())
    return OPENING.replace("NORMAL", normal) + body + CLOSING


def kept(path: Path) -> dict[str, dict[str, str]]:
    """Each case sheet's first row and columns as the saved file keeps them, by the case's name."""
    out: dict[str, dict[str, str]] = {}
    with zipfile.ZipFile(path) as package:
        for index, name in enumerate(CASES, start=2):
            sheet = package.read(f"xl/worksheets/sheet{index}.xml").decode("utf-8")
            row = re.search(r'<row r="1"[^>]*>', sheet)
            columns = re.search(r"<cols>.*?</cols>", sheet)
            out[name] = {"row": row.group(0) if row else "", "cols": columns.group(0) if columns else ""}
    return out


def main() -> None:
    probes: list[dict[str, object]] = []
    with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        standard = excel.run_vba("Function H() As String\nH = CStr(ActiveSheet.StandardHeight)\nEnd Function", "H")
        assert standard.value == "15", f"Excel lays sheets out at another DPI: StandardHeight {standard.value}"
        version = excel.run_vba('Function V() As String\nV = Application.Version & "|" & Application.Build\n'
                                'End Function', "V")
        for normal in NORMALS:
            path = Path(folder) / f"{normal.replace(' ', '_')}.xlsx"
            source = module(normal)
            result = excel.run_vba(source, "Probe", [str(path)], timeout=300.0)
            assert result.ok and not result.dialogs, f"{result.outcome}: {result.message} {result.error}"
            print(normal, result.value, flush=True)
            probes.append({"normal": normal, "macro": source, "result": result.value, "file": kept(path)})
    OUT.write_text(json.dumps({"excel_version_build": version.value,
                               "measured_at": datetime.now(timezone.utc).isoformat(),
                               "cases": CASES, "probes": probes}, indent=2) + "\n", encoding="utf-8")


if __name__ == "__main__":
    main()
