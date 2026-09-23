"""How Excel writes a number format to its file, and reads one back from it.

Range.NumberFormat and the file spell a format differently: Excel reads
"$#,##0_);[Red]($#,##0)" but writes numFmtId 6 with the code
"$"#,##0_);[Red]\\("$"#,##0\\) spelled out, while "0%" is numFmtId 9 with
nothing written. The probe sets every built-in format and a set of
custom ones on cells of their own, reads NumberFormat back and saves the
workbook, so the file shows each format's id and code. It then plants
codes in the saved stylesheet, spelt the ways another program might
spell them, opens that file and reads what NumberFormat makes of each.

    python scripts/measure_format_codes.py

writes tests/fixtures/format_codes/: format_codes.xlsx as Excel saved it,
planted_source.xlsx as planted, and format_codes.json with the codes,
what NumberFormat read, and what each planted code read as;
tests/test_excel_value_typing.py replays it.
"""

from __future__ import annotations

import json
import re
import shutil
import tempfile
import zipfile
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "tests" / "fixtures" / "format_codes"

#: Every built-in format as NumberFormat spells it, then custom ones.
CODES = [
    "General", "0", "0.00", "#,##0", "#,##0.00", "$#,##0_);($#,##0)", "$#,##0_);[Red]($#,##0)",
    "$#,##0.00_);($#,##0.00)", "$#,##0.00_);[Red]($#,##0.00)", "0%", "0.00%", "0.00E+00", "# ?/?", "# ??/??",
    "m/d/yyyy", "d-mmm-yy", "d-mmm", "mmm-yy", "h:mm AM/PM", "h:mm:ss AM/PM", "h:mm", "h:mm:ss", "m/d/yyyy h:mm",
    "#,##0_);(#,##0)", "#,##0_);[Red](#,##0)", "#,##0.00_);(#,##0.00)", "#,##0.00_);[Red](#,##0.00)",
    '_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)', '_($* #,##0_);_($* (#,##0);_($* "-"_);_(@_)',
    '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)', '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)',
    "mm:ss", "[h]:mm:ss", "mm:ss.0", "##0.0E+0", "@",
    "$#,##0.00", "$#,##0", "0.0", "0.000%", "yyyy-mm-dd", "dddd", "[$-409]h:mm:ss AM/PM", '0 "units"',
    "#,##0 [$USD]", "[Red]0.00", "0.00;-0.00", "[h]:mm", "d", "mmmm", "hh", "0.00_);(0.00)", "# ?/8", "yyyy",
    "[$-409]m/d/yyyy", "m/d/yy", "mm/dd/yyyy", "ss", "[$$-409]#,##0.00", "0.00%;[Red]-0.00%", "0-0", '0" "0',
    '"abc"0', "\\$0", "\\@", '"@"', "0;-0;;@", "(0)", "0_)", "0 %", "[>100]0;0.00", "[Blue][<0]0;0",
    "#,##0;(#,##0)", "$0.00;($0.00)", '0.00 "$"', "h:mm:ss.000", "[mm]:ss", "[ss]", "d/m/yyyy", "0.00\\%",
    '#,##0.00_);[Red]\\(#,##0.00\\)', "0,", "0.0,,", "#,##0,\"K\"", "?/?", "0.00;;;", ";;;", "@\" x\"",
]

#: Codes planted in the saved stylesheet under ids of their own, spelt as another program might spell them.
PLANTED = [
    '"$"#,##0.00', "$#,##0.00", "\\$#,##0.00", "#,##0.00_);\\(#,##0.00\\)", "#,##0.00_);(#,##0.00)",
    '"$"#,##0_);[Red]\\("$"#,##0\\)', "0\\-0", "0-0", '0"-"0', "\\(0\\)", "(0)", "[red]0.00", "general",
    "M/D/YYYY", "mmss.0", "mm:ss.0", '"@"', "0.00\\%", "0.00%",
]
#: Built-in ids planted with no numFmt of their own, which Excel has to know.
PLANTED_IDS = [5, 6, 7, 8, 12, 13, 37, 38, 39, 40, 41, 42, 43, 44, 45, 46, 47, 48, 49]


def vba_text(text: str) -> str:
    return '"' + text.replace('"', '""') + '"'


def run(excel: ExcelSession, code: str, procedure: str) -> str:
    result = excel.run_vba(code, procedure, timeout=600.0)
    assert result.ok, f"{result.outcome}: {result.message} {result.error}"
    return str(result.value)


def plant(source: Path, target: Path) -> None:
    """The saved workbook with one cell per planted code and per bare built-in id, below the measured ones."""
    with zipfile.ZipFile(source) as package:
        parts = {name: package.read(name) for name in package.namelist()}
    styles = parts["xl/styles.xml"].decode("utf-8")
    sheet = parts["xl/worksheets/sheet1.xml"].decode("utf-8")
    formats = [f'<numFmt numFmtId="{300 + index}" formatCode="{escape(code)}"/>' for index, code in enumerate(PLANTED)]
    if "<numFmts" in styles:
        styles = re.sub(r"</numFmts>", "".join(formats) + "</numFmts>", styles, count=1)
    else:
        styles = styles.replace("<fonts", f'<numFmts count="{len(formats)}">{"".join(formats)}</numFmts><fonts', 1)
    styles = re.sub(r'(<numFmts count=")\d+', lambda m: m.group(1) + str(styles.count("<numFmt ")), styles, count=1)
    ids = [300 + index for index in range(len(PLANTED))] + PLANTED_IDS
    xfs = "".join(f'<xf numFmtId="{one}" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
                  for one in ids)
    first = int(re.search(r'<cellXfs count="(\d+)"', styles).group(1))  # type: ignore[union-attr]
    styles = styles.replace("</cellXfs>", xfs + "</cellXfs>", 1)
    styles = re.sub(r'<cellXfs count="\d+"', f'<cellXfs count="{first + len(ids)}"', styles, count=1)
    top = len(CODES) + 1
    rows = "".join(f'<row r="{top + index}"><c r="A{top + index}" s="{first + index}"><v>1</v></c></row>'
                   for index in range(len(ids)))
    sheet = sheet.replace("</sheetData>", rows + "</sheetData>", 1)
    sheet = re.sub(r'<dimension ref="[^"]*"/>', f'<dimension ref="A1:A{top + len(ids) - 1}"/>', sheet, count=1)
    parts["xl/styles.xml"], parts["xl/worksheets/sheet1.xml"] = styles.encode("utf-8"), sheet.encode("utf-8")
    with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as package:
        for name, data in parts.items():
            package.writestr(name, data)


def escape(text: str) -> str:
    return text.replace("&", "&amp;").replace('"', "&quot;").replace("<", "&lt;").replace(">", "&gt;")


def main() -> None:
    folder = Path(tempfile.mkdtemp())
    saved, planted = folder / "format_codes.xlsx", folder / "planted_source.xlsx"
    build = ["Public Function Build() As String", "Dim wb As Object, ws As Object, out As String",
             "Application.DisplayAlerts = False", "Set wb = Workbooks.Add(xlWBATWorksheet)", "Set ws = wb.Worksheets(1)",
             "On Error Resume Next"]
    for row, code in enumerate(CODES, start=1):
        build += ["Err.Clear", f"ws.Cells({row}, 1).Value = 1", f"ws.Cells({row}, 1).NumberFormat = {vba_text(code)}",
                  f'If Err.Number <> 0 Then out = out & "S" & Err.Number & "^" Else out = out & '
                  f'ws.Cells({row}, 1).NumberFormat & "^"']
    build += ["On Error GoTo 0", f'wb.SaveAs Filename:="{saved}", FileFormat:=51', "wb.Close False", "Build = out",
              "End Function"]
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        read = run(excel, "\n".join(build) + "\n", "Build").split("^")[: len(CODES)]
        plant(saved, planted)
        count = len(PLANTED) + len(PLANTED_IDS)
        reopen = ["Public Function Reopen() As String", "Dim wb As Object, out As String, i As Long",
                  "Application.DisplayAlerts = False", f'Set wb = Workbooks.Open("{planted}")',
                  f"For i = {len(CODES) + 1} To {len(CODES) + count}",
                  'out = out & wb.Worksheets(1).Cells(i, 1).NumberFormat & "^"', "Next", "wb.Close False",
                  "Reopen = out", "End Function"]
        planted_read = run(excel, "\n".join(reopen) + "\n", "Reopen").split("^")[:count]
    OUT.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(saved, OUT / "format_codes.xlsx")
    shutil.copyfile(planted, OUT / "planted_source.xlsx")
    record = {"codes": [{"code": code, "read": answer} for code, answer in zip(CODES, read, strict=True)],
              "planted": [{"code": code, "read": answer}
                          for code, answer in zip(PLANTED, planted_read[: len(PLANTED)], strict=True)],
              "planted_ids": [{"id": one, "read": answer}
                              for one, answer in zip(PLANTED_IDS, planted_read[len(PLANTED):], strict=True)]}
    (OUT / "format_codes.json").write_text(json.dumps(record, indent=1) + "\n", encoding="utf-8")
    for entry in record["codes"]:
        print(f"{entry['code']:52} {entry['read']}")
    for entry in record["planted"]:
        print(f"planted {entry['code']:44} {entry['read']}")
    for entry in record["planted_ids"]:
        print(f"planted id {entry['id']:<5} {entry['read']}")


if __name__ == "__main__":
    main()
