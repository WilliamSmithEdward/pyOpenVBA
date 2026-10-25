"""A worksheet's tab colour, Worksheet.Tab, in live Excel.

A macro adds a sheet for each case to one new workbook and sets its tab's
Color, ColorIndex, ThemeColor, TintAndShade or Pattern, one after another
for some, an error's number noted where a setting raises one. It reads
back every one of them with its type. One sheet is copied. The workbook
is saved, each sheet's sheetPr kept, and opened again, and every tab is
read once more. Then a tab is given each of 1,397 colours, random ones,
near-greys, very light and dark ones and greys about the points halfway
between the palette's greys, and the ColorIndex it reads back kept.

    python scripts/measure_sheet_tabs.py

writes tests/fixtures/sheet_tabs/: the workbook as Excel saved it and
sheet_tabs.json, which tests/test_excel_sheet_tabs.py replays.
"""

from __future__ import annotations

import json
import random
import re
import shutil
import sys
import zipfile
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "scripts"))
from fixture_workbook import strip_save_path  # noqa: E402

OUT = ROOT / "tests" / "fixtures" / "sheet_tabs"

#: name -> the settings made, in order: (the Tab's property, a VBA expression for its value).
CASES: dict[str, list[tuple[str, str]]] = {
    "none": [],
    "rgb_red": [("Color", "RGB(255, 0, 0)")],
    "rgb_odd": [("Color", "RGB(12, 34, 56)")],
    "rgb_black": [("Color", "0")],
    "rgb_white": [("Color", "RGB(255, 255, 255)")],
    "rgb_double": [("Color", "255.6")],
    "rgb_text": [("Color", '"255"')],
    "rgb_negative": [("Color", "-1")],
    "rgb_too_big": [("Color", "16777216")],
    "color_none": [("Color", "RGB(255, 0, 0)"), ("Color", "xlNone")],
    "color_false": [("Color", "RGB(255, 0, 0)"), ("Color", "False")],
    "color_automatic": [("Color", "RGB(255, 0, 0)"), ("Color", "xlAutomatic")],
    "index_3": [("ColorIndex", "3")],
    "index_5": [("ColorIndex", "5")],
    "index_56": [("ColorIndex", "56")],
    "index_57": [("ColorIndex", "57")],
    "index_0": [("ColorIndex", "0")],
    "index_none": [("Color", "RGB(255, 0, 0)"), ("ColorIndex", "xlColorIndexNone")],
    "index_automatic": [("ColorIndex", "xlColorIndexAutomatic")],
    "index_text": [("ColorIndex", '"7"')],
    "theme_dark1": [("ThemeColor", "xlThemeColorDark1")],
    "theme_light1": [("ThemeColor", "xlThemeColorLight1")],
    "theme_dark2": [("ThemeColor", "xlThemeColorDark2")],
    "theme_light2": [("ThemeColor", "xlThemeColorLight2")],
    "theme_accent1": [("ThemeColor", "xlThemeColorAccent1")],
    "theme_accent6": [("ThemeColor", "xlThemeColorAccent6")],
    "theme_hyperlink": [("ThemeColor", "xlThemeColorHyperlink")],
    "theme_followed": [("ThemeColor", "xlThemeColorFollowedHyperlink")],
    "theme_0": [("ThemeColor", "0")],
    "theme_13": [("ThemeColor", "13")],
    "theme_tint_up": [("ThemeColor", "xlThemeColorAccent2"), ("TintAndShade", "0.4")],
    "theme_tint_down": [("ThemeColor", "xlThemeColorAccent2"), ("TintAndShade", "-0.25")],
    "theme_tint_full": [("ThemeColor", "xlThemeColorAccent4"), ("TintAndShade", "1")],
    "theme_tint_beyond": [("ThemeColor", "xlThemeColorAccent4"), ("TintAndShade", "1.5")],
    "theme_tint_then_theme": [("ThemeColor", "xlThemeColorAccent2"), ("TintAndShade", "0.4"),
                              ("ThemeColor", "xlThemeColorAccent5")],
    "tint_alone": [("TintAndShade", "0.5")],
    "tint_then_theme": [("TintAndShade", "0.5"), ("ThemeColor", "xlThemeColorAccent3")],
    "rgb_tint": [("Color", "RGB(255, 0, 0)"), ("TintAndShade", "0.5")],
    "index_tint": [("ColorIndex", "5"), ("TintAndShade", "-0.5")],
    "index_then_theme": [("ColorIndex", "3"), ("ThemeColor", "xlThemeColorAccent6")],
    "theme_then_rgb": [("ThemeColor", "xlThemeColorAccent1"), ("Color", "RGB(0, 255, 0)")],
    "theme_then_index": [("ThemeColor", "xlThemeColorAccent1"), ("ColorIndex", "10")],
    "pattern_solid": [("Color", "RGB(255, 0, 0)"), ("Pattern", "xlSolid")],
    "pattern_none": [("Color", "RGB(255, 0, 0)"), ("Pattern", "xlNone")],
    "pattern_gray": [("Color", "RGB(255, 0, 0)"), ("Pattern", "xlGray50")],
    "pattern_alone": [("Pattern", "xlSolid")],
}

HELPER = '''Private Function Got(target As Object, name As String) As String
    Dim v As Variant
    On Error GoTo Failed
    v = CallByName(target, name, VbGet)
    Got = TypeName(v) & ":" & CStr(v)
    Exit Function
Failed:
    Got = "E" & Err.Number
End Function

Private Function Dump(ws As Object) As String
    Dim t As Object
    Set t = ws.Tab
    Dump = TypeName(t) & "|" & Got(t, "Color") & "|" & Got(t, "ColorIndex") & "|" & Got(t, "ThemeColor") & "|" & _
        Got(t, "TintAndShade") & "|" & Got(t, "Pattern")
End Function

Private Function Try(ws As Object, name As String, value As Variant) As String
    On Error GoTo Failed
    CallByName ws.Tab, name, VbLet, value
    Exit Function
Failed:
    Try = "E" & Err.Number & ":" & name & " "
End Function
'''

BUILD = HELPER + '''
Public Function Build(folder As String) As String
    Dim wb As Object, ws As Object, out As String, notes As String, name As String
    Application.DisplayAlerts = False
    Set wb = Workbooks.Add(xlWBATWorksheet)
{cases}
    wb.Worksheets("rgb_red").Copy After:=wb.Worksheets(wb.Worksheets.Count)
    out = out & "copy^" & Dump(wb.Worksheets(wb.Worksheets.Count)) & "~"
    wb.Worksheets(1).Activate
    name = "sheet_tabs.xlsx"
    wb.SaveAs folder & "\\" & name, 51
    wb.Close False
    Set wb = Workbooks.Open(folder & "\\" & name)
    out = out & "@@"
    For Each ws In wb.Worksheets
        out = out & ws.Name & "^" & Dump(ws) & "~"
    Next
    wb.Close False
    Build = out
End Function
'''


NEAREST = '''Public Function Nearest(list As String) As String
    Dim wb As Object, ws As Object, part As Variant, out As String
    Application.DisplayAlerts = False
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Worksheets(1)
    For Each part In Split(list, ",")
        ws.Tab.Color = CLng(part)
        out = out & ws.Tab.ColorIndex & ","
    Next
    wb.Close False
    Nearest = out
End Function
'''


def _clamp(value: int) -> int:
    return max(0, min(255, value))


def nearest_colours() -> list[tuple[int, int, int]]:
    """The colours whose ColorIndex the tab is read for: a seeded spread of random colours, near-greys, very light
    and dark colours, the palette greys' halfway points and chroma 10 against 11."""
    pick = random.Random(20261025)
    colours = [(12, 34, 56), (247, 239, 255), (70, 120, 134), (150, 96, 125), (241, 169, 131), (190, 80, 20),
               (25, 107, 36), (128, 128, 128), (200, 200, 200), (10, 10, 10)]
    colours += [(pick.randrange(256), pick.randrange(256), pick.randrange(256)) for _ in range(540)]
    for _ in range(150):
        grey = pick.randrange(256)
        colours.append((_clamp(grey + pick.randint(-14, 14)), _clamp(grey + pick.randint(-14, 14)),
                        _clamp(grey + pick.randint(-14, 14))))
    for _ in range(100):
        base = pick.choice((pick.randrange(0, 40), pick.randrange(215, 256)))
        colours.append((_clamp(base + pick.randint(-25, 25)), _clamp(base + pick.randint(-25, 25)),
                        _clamp(base + pick.randint(-25, 25))))
    colours += [(grey, grey, grey) for grey in range(0, 256, 5)]
    for middle in (25, 89, 139, 171, 223):
        for grey in range(middle - 8, middle + 9):
            colours.append((grey, grey, grey))
            for _ in range(4):
                spread = pick.randint(1, 10)
                low = _clamp(grey - pick.randint(0, spread))
                channels = [low, _clamp(low + spread), _clamp(low + pick.randint(0, spread))]
                pick.shuffle(channels)
                colours.append((channels[0], channels[1], channels[2]))
    for base in range(10, 250, 12):
        for spread in (10, 11):
            for order in ((0, 1, 2), (2, 0, 1), (1, 2, 0)):
                channels = [base, base + spread, base + spread // 2]
                colours.append((_clamp(channels[order[0]]), _clamp(channels[order[1]]), _clamp(channels[order[2]])))
    return colours


def case_code(index: int, name: str, steps: list[tuple[str, str]]) -> str:
    lines = ["    Set ws = wb.Worksheets(1)" if index == 0 else
             "    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))",
             f'    ws.Name = "{name}"', '    notes = ""']
    lines += [f'    notes = notes & Try(ws, "{prop}", {value})' for prop, value in steps]
    lines.append(f'    out = out & "{name}^" & notes & Dump(ws) & "~"')
    return "\n".join(lines) + "\n"


def sheet_properties(path: Path) -> dict[str, str]:
    """Each sheet's sheetPr as the file keeps it, by the sheet's name."""
    with zipfile.ZipFile(path) as package:
        workbook = package.read("xl/workbook.xml").decode("utf-8")
        rels = package.read("xl/_rels/workbook.xml.rels").decode("utf-8")
        targets = {found.group(1): found.group(2) for found in
                   re.finditer(r'<Relationship\b[^>]*\bId="([^"]*)"[^>]*\bTarget="([^"]*)"', rels)}
        out: dict[str, str] = {}
        for name, relationship in re.findall(r'<sheet\b[^>]*\bname="([^"]*)"[^>]*\br:id="([^"]*)"', workbook):
            part = "xl/" + targets[relationship].lstrip("/").removeprefix("xl/")
            text = package.read(part).decode("utf-8")
            found = re.search(r"<sheetPr\b[^>]*/>|<sheetPr\b.*?</sheetPr>", text, re.DOTALL)
            out[name] = found.group() if found else ""
        return out


def main() -> None:
    if OUT.exists():
        shutil.rmtree(OUT)
    OUT.mkdir(parents=True)
    cases = "".join(case_code(index, name, steps) for index, (name, steps) in enumerate(CASES.items()))
    code = BUILD.replace("{cases}", cases)
    colours = nearest_colours()
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(code, "Build", (str(OUT),), timeout=600.0)
        assert result.ok, f"{result.outcome}: {result.message} {result.error}"
        listed = ",".join(str(red + green * 256 + blue * 65536) for red, green, blue in colours)
        read = excel.run_vba(NEAREST, "Nearest", (listed,), timeout=600.0)
        assert read.ok, f"{read.outcome}: {read.message} {read.error}"
    indexes = [int(part) for part in str(read.value).split(",") if part]
    session, reopened = str(result.value).split("@@", 1)
    path = OUT / "sheet_tabs.xlsx"
    strip_save_path(path)
    record = {"build": code, "cases": [{"name": name, "steps": steps} for name, steps in CASES.items()],
              "session": dict(one.split("^", 1) for one in session.split("~") if "^" in one),
              "reopened": dict(one.split("^", 1) for one in reopened.split("~") if "^" in one),
              "sheet_properties": sheet_properties(path),
              "nearest": [[list(colour), index] for colour, index in zip(colours, indexes, strict=True)]}
    (OUT / "sheet_tabs.json").write_text(json.dumps(record, indent=1, ensure_ascii=False) + "\n", encoding="utf-8")
    for name, dump in record["session"].items():
        print(f"{name:22} {dump:90} {record['sheet_properties'].get(name, '')}")


if __name__ == "__main__":
    main()
