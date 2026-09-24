"""An Excel-authored workbook of formatted cells, with what Excel reads back from it.

Excel formats one cell per case in a new workbook, saves it as .xlsx,
closes it, opens it again and records every formatting property of every
case. The workbook and the answers are the load fixture: the in-memory
model must read the same values from the same file. Writing the same
cases from the model and asking Excel for them is the live gate.

    python scripts/measure_format_file.py

writes tests/fixtures/format/formats.xlsx and formats_answers.json.
"""

from __future__ import annotations

import json
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

from fixture_workbook import strip_save_path

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "tests" / "fixtures" / "format"

#: One formatting case per row of column A. `c` is the cell being formatted.
CASES: dict[str, str] = {
    "default": "",
    "bold": "c.Font.Bold = True",
    "italic_underline": "c.Font.Italic = True\nc.Font.Underline = xlUnderlineStyleSingle",
    "name_size": 'c.Font.Name = "Arial"\nc.Font.Size = 14',
    "font_rgb": "c.Font.Color = RGB(255, 0, 0)",
    "font_theme_tint": "c.Font.ThemeColor = xlThemeColorAccent1\nc.Font.TintAndShade = 0.4",
    "strike_superscript": "c.Font.Strikethrough = True\nc.Font.Superscript = True",
    "subscript": "c.Font.Subscript = True",
    "double_accounting": "c.Font.Underline = xlUnderlineStyleDoubleAccounting",
    "font_color_index": "c.Font.ColorIndex = 45",
    "font_automatic": "c.Font.ColorIndex = xlColorIndexAutomatic",
    "font_half_size": "c.Font.Size = 10.5",
    "font_name_back": 'c.Font.Name = "Arial"\nc.Font.Name = "Aptos Narrow"\nc.Font.Bold = True',
    "font_tint_automatic": "c.Font.ColorIndex = xlColorIndexAutomatic\nc.Font.TintAndShade = 0.5",
    "font_tint_indexed": "c.Font.ColorIndex = 3\nc.Font.TintAndShade = 0.5",
    "fill_tint_indexed": "c.Interior.ColorIndex = 6\nc.Interior.TintAndShade = 0.5",
    "fill_pattern_color_first": "c.Interior.PatternColor = RGB(0, 0, 255)",
    "fill_pattern_automatic": "c.Interior.Pattern = xlPatternAutomatic",
    "border_tint_indexed": "c.Borders(xlEdgeLeft).ColorIndex = 10\nc.Borders(xlEdgeLeft).TintAndShade = 0.5",
    "font_major": 'c.Font.Name = "Aptos Display"',
    # Pairs whose colour differs between the tint as set and the n/32767 tint the file stores.
    "tint_black_tenth": "c.Font.ThemeColor = xlThemeColorLight1\nc.Font.TintAndShade = 0.1",
    "tint_black_quarter": "c.Font.ThemeColor = xlThemeColorLight1\nc.Font.TintAndShade = 0.25",
    "tint_accent6_tenth": "c.Interior.ThemeColor = xlThemeColorAccent6\nc.Interior.TintAndShade = 0.1",
    "tint_accent5_quarter": "c.Interior.ThemeColor = xlThemeColorAccent5\nc.Interior.TintAndShade = 0.25",
    "tint_hyperlink_border": "c.Borders(xlEdgeLeft).ThemeColor = xlThemeColorHyperlink\n"
                             "c.Borders(xlEdgeLeft).TintAndShade = 0.25",
    "everything_font": 'c.Font.Name = "Courier New"\nc.Font.Size = 9\nc.Font.Bold = True\nc.Font.Italic = True\n'
                       "c.Font.Strikethrough = True\nc.Font.Underline = xlUnderlineStyleDouble\n"
                       "c.Font.Color = RGB(0, 112, 192)",
    "fill_rgb": "c.Interior.Color = RGB(255, 255, 0)",
    "fill_theme_tint": "c.Interior.ThemeColor = xlThemeColorAccent2\nc.Interior.TintAndShade = 0.6",
    "fill_color_index": "c.Interior.ColorIndex = 45",
    "fill_automatic": "c.Interior.ColorIndex = xlColorIndexAutomatic",
    "fill_pattern_colors": "c.Interior.Pattern = xlChecker\nc.Interior.PatternColor = RGB(0, 0, 255)\n"
                           "c.Interior.Color = RGB(0, 255, 0)",
    "fill_gray": "c.Interior.Pattern = xlGray50",
    "fill_solid_then_pattern": "c.Interior.Color = RGB(255, 255, 0)\nc.Interior.Pattern = xlChecker",
    "fill_pattern_theme": "c.Interior.Pattern = xlChecker\nc.Interior.PatternThemeColor = xlThemeColorAccent4\n"
                          "c.Interior.PatternTintAndShade = 0.25",
    "border_thin_left": "c.Borders(xlEdgeLeft).LineStyle = xlContinuous",
    "border_around_red": "c.BorderAround xlContinuous, xlMedium, , RGB(255, 0, 0)",
    "border_double_bottom": "c.Borders(xlEdgeBottom).LineStyle = xlDouble",
    "border_diagonal_up": "c.Borders(xlDiagonalUp).LineStyle = xlDashDot",
    "border_diagonal_both": "c.Borders(xlDiagonalUp).LineStyle = xlContinuous\n"
                            "c.Borders(xlDiagonalDown).LineStyle = xlContinuous",
    "border_theme_tint": "c.Borders(xlEdgeLeft).ThemeColor = xlThemeColorAccent3\n"
                        "c.Borders(xlEdgeLeft).TintAndShade = -0.5",
    "border_slant": "c.Borders(xlEdgeTop).LineStyle = xlSlantDashDot",
    "border_hairline": "c.Borders(xlEdgeRight).Weight = xlHairline",
    "border_medium_dash": "c.Borders(xlEdgeRight).LineStyle = xlDash\nc.Borders(xlEdgeRight).Weight = xlMedium",
    "border_dot": "c.Borders(xlEdgeTop).LineStyle = xlDot",
    "border_color_index": "c.Borders(xlEdgeBottom).ColorIndex = 10",
    "align_center_top_wrap": "c.HorizontalAlignment = xlCenter\nc.VerticalAlignment = xlTop\nc.WrapText = True",
    "align_indent_right": "c.HorizontalAlignment = xlRight\nc.IndentLevel = 3",
    "align_orientation": "c.Orientation = 45",
    "align_vertical_text": "c.Orientation = xlVertical",
    "align_upward": "c.Orientation = xlUpward",
    "align_shrink": "c.ShrinkToFit = True",
    "align_rtl": "c.ReadingOrder = xlRTL",
    "align_distributed_indent": "c.HorizontalAlignment = xlDistributed\nc.AddIndent = True\nc.IndentLevel = 1",
    "align_justify_vertical": "c.VerticalAlignment = xlJustify",
    "protect": "c.Locked = False\nc.FormulaHidden = True",
    "number_format": 'c.Value = 0.5\nc.NumberFormat = "0.00%"',
    "mixed_everything": 'c.Value = "x"\nc.Font.Bold = True\nc.Interior.Color = RGB(200, 230, 255)\n'
                        "c.Borders.LineStyle = xlContinuous\nc.HorizontalAlignment = xlCenter\nc.Locked = False",
}

FONT = ("Name", "Size", "Bold", "Italic", "Underline", "Strikethrough", "Color", "ColorIndex",
        "ThemeColor", "TintAndShade", "Superscript", "Subscript", "FontStyle")
FILL = ("Color", "ColorIndex", "Pattern", "PatternColor", "PatternColorIndex", "ThemeColor",
        "TintAndShade", "PatternThemeColor", "PatternTintAndShade")
CELL = ("HorizontalAlignment", "VerticalAlignment", "WrapText", "IndentLevel", "Orientation",
        "ShrinkToFit", "AddIndent", "ReadingOrder", "Locked", "FormulaHidden", "NumberFormat")
BORDER = ("LineStyle", "Weight", "Color", "ColorIndex", "ThemeColor", "TintAndShade")


def reads() -> list[str]:
    found = [f"Font.{name}" for name in FONT] + [f"Interior.{name}" for name in FILL] + list(CELL)
    found += [f"Borders({index}).{name}" for index in (5, 6, 7, 8, 9, 10) for name in BORDER]
    return found


READS = reads()

HELPER = '''Private Function Show(v As Variant) As String
    If IsNull(v) Then
        Show = "Null"
    ElseIf IsEmpty(v) Then
        Show = "Empty"
    Else
        Show = TypeName(v) & ":" & CStr(v)
    End If
End Function
'''


def describe_code() -> str:
    """VBA that reports every read of one cell, the same code the replay test runs."""
    lines = ["Public Function Describe(c As Object) As String", "Dim v As Variant", "On Error Resume Next"]
    for expression in READS:
        lines += ["Err.Clear", "v = Empty", f"v = c.{expression}",
                  'If Err.Number <> 0 Then Describe = Describe & "E" & Err.Number & ";" '
                  'Else Describe = Describe & Show(v) & ";"']
    lines.append("End Function")
    return HELPER + "\n".join(lines) + "\n"


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    target = OUT / "formats.xlsx"
    if target.exists():
        target.unlink()
    build = ["Public Function Build() As String", "Dim wb As Object, c As Object",
             "Application.DisplayAlerts = False", "Set wb = Workbooks.Add(xlWBATWorksheet)"]
    for row, (name, setup) in enumerate(CASES.items(), start=1):
        build.append(f'Set c = wb.Worksheets(1).Range("A{row}")')
        build.append(f'wb.Worksheets(1).Range("B{row}").Value = "{name}"')
        build += [line.strip() for line in setup.splitlines()]
    build += [f'wb.SaveAs Filename:="{target}", FileFormat:=51', "wb.Close False",
              f'Set wb = Workbooks.Open("{target}")', "For Each c In wb.Worksheets(1).Range(\"A1:A"
              f'{len(CASES)}")', 'Build = Build & Describe(c) & "|"', "Next c", "wb.Close False",
              "End Function"]
    code = describe_code() + "\n".join(build) + "\n"
    with ExcelSession(HarnessConfig(lock_wait_s=45.0)) as excel:
        excel.new_document()
        result = excel.run_vba(code, "Build", timeout=240.0)
        assert result.ok, f"{result.outcome}: {result.message}"
    strip_save_path(target)
    described = str(result.value).split("|")[:-1]
    answers = {name: dict(zip(READS, text.split(";")[:-1], strict=True))
               for name, text in zip(CASES, described, strict=True)}
    payload = {"reads": READS, "describe": describe_code(), "cases": list(CASES), "setups": CASES,
               "answers": answers}
    (OUT / "formats_answers.json").write_text(json.dumps(payload, indent=1) + "\n", encoding="utf-8")
    print(f"{len(answers)} cases, {len(READS)} reads each -> {OUT}")


if __name__ == "__main__":
    main()
