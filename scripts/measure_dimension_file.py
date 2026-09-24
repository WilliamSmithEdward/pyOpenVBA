"""Row heights, column widths and hidden rows and columns as Excel saves them and reads them back.

Three workbooks, all written to tests/fixtures/dimensions/:

- dimensions.xlsx: every case on a sheet of its own, set by Excel and
  saved. dimensions_answers.json holds each case's setup, what Excel reads
  from the sheet after opening the file again, and the parts of the sheet
  XML that say how big things are: sheetFormatPr, cols and every row's
  start tag.
- foreign.xlsx: sizes planted in the XML the way another program might
  write them (heights without customHeight, widths off Excel's pixel
  grid, negative and oversized values), with what Excel reads from them
  and how it writes them back when it saves.
- template_resaved.xlsx: pyOpenVBA's own new-workbook template, opened by
  Excel on this display, given one value and one row height, and saved,
  which shows what Excel rewrites in a sheet that came from a 144-DPI
  display.

    python scripts/measure_dimension_file.py
"""

from __future__ import annotations

import json
import re
import sys
import tempfile
import zipfile
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

from fixture_workbook import strip_save_path

ROOT = Path(__file__).resolve().parent.parent
FIXTURES = ROOT / "tests" / "fixtures" / "dimensions"
sys.path.insert(0, str(ROOT / "src"))

ROWS = 12
COLUMNS = 12

#: Describe(ws) reads these for every sheet, then the per-row and per-column answers.
DESCRIBE = f'''Public Function Describe(ws As Object) As String
    Dim out As String, r As Long, c As Long
    On Error Resume Next
    out = ws.StandardHeight & ";" & ws.StandardWidth & ";" & ws.UsedRange.Address & ";"
    For r = 1 To {ROWS}
        out = out & ws.Rows(r).RowHeight & ":" & ws.Rows(r).Height & ":" & ws.Rows(r).Hidden & ":" & _
            ws.Rows(r).UseStandardHeight & ";"
    Next r
    For c = 1 To {COLUMNS}
        out = out & ws.Columns(c).ColumnWidth & ":" & ws.Columns(c).Width & ":" & ws.Columns(c).Hidden & ":" & _
            ws.Columns(c).UseStandardWidth & ";"
    Next c
    Describe = out
End Function
'''

CASES: dict[str, str] = {
    "custom_rows": "ws.Rows(2).RowHeight = 20\nws.Rows(3).RowHeight = 0.3\nws.Rows(4).RowHeight = 409.5\n"
                   "ws.Rows(5).RowHeight = 15\nws.Rows(6).RowHeight = 12.3\nws.Rows(7).RowHeight = 1.15\n"
                   "ws.Rows(8).RowHeight = 19.9",
    "hidden_rows": "ws.Rows(2).RowHeight = 20\nws.Rows(2).Hidden = True\nws.Rows(3).Hidden = True\n"
                   "ws.Rows(4).RowHeight = 0\nws.Rows(5).RowHeight = 20\nws.Rows(5).RowHeight = 0\n"
                   "ws.Rows(6).RowHeight = 0.1\nws.Rows(7).RowHeight = 0.05",
    "custom_columns": "ws.Columns(2).ColumnWidth = 20\nws.Columns(3).ColumnWidth = 0.5\nws.Columns(4).ColumnWidth = 255\n"
                      "ws.Columns(5).ColumnWidth = 8.43\nws.Columns(6).ColumnWidth = 10.1\n"
                      'ws.Columns("G:I").ColumnWidth = 12\nws.Columns(11).ColumnWidth = 12',
    "hidden_columns": "ws.Columns(2).ColumnWidth = 20\nws.Columns(2).Hidden = True\nws.Columns(3).Hidden = True\n"
                      "ws.Columns(4).ColumnWidth = 0\nws.Columns(5).Hidden = True\nws.Columns(5).Hidden = False\n"
                      "ws.Columns(6).ColumnWidth = 0.04\nws.Columns(7).ColumnWidth = 20\nws.Columns(7).ColumnWidth = 0",
    "standard_width": "ws.StandardWidth = 12\nws.Columns(2).ColumnWidth = 20",
    "standard_width_zero": "ws.StandardWidth = 0",
    "standard_width_back": "ws.StandardWidth = 12\nws.StandardWidth = 8.43",
    "use_standard": "ws.Rows(2).RowHeight = 20\nws.Rows(2).UseStandardHeight = True\nws.Rows(3).UseStandardHeight = False\n"
                    "ws.Columns(2).ColumnWidth = 20\nws.Columns(2).UseStandardWidth = True\nws.Rows(4).RowHeight = 15",
    "all_rows": 'ws.Rows("1:1048576").RowHeight = 20',
    "all_columns": 'ws.Columns("A:XFD").ColumnWidth = 12',
    "cells_and_rows": 'ws.Range("A2").Value = 1\nws.Range("C2").Value = 2\nws.Range("B20").Value = 3\n'
                      "ws.Rows(2).RowHeight = 20\nws.Rows(3).RowHeight = 30\nws.Rows(18).Hidden = True\n"
                      'ws.Range("E5").Font.Bold = True',
    "insert_rows": "ws.Rows(3).RowHeight = 30\nws.Rows(5).Hidden = True\nws.Rows(4).Insert",
    # The same, with the used range read before the save: Excel then drops
    # the customFormat="1" it leaves on rows after an insert.
    "insert_rows_read": "ws.Rows(3).RowHeight = 30\nws.Rows(5).Hidden = True\nws.Rows(4).Insert\n"
                        "Debug.Print ws.UsedRange.Address",
    "delete_columns": 'ws.Columns(3).ColumnWidth = 30\nws.Columns(6).Hidden = True\nws.Columns(4).ColumnWidth = 40\n'
                      'ws.Columns("D:E").Delete',
    "autofit": "ws.Rows(2).RowHeight = 30\nws.Rows(3).RowHeight = 30\nws.Rows(3).Hidden = True\nws.Rows(\"2:3\").AutoFit",
    "hide_every_row": 'ws.Rows("1:1048576").Hidden = True',
    "show_one_of_every_row": 'ws.Rows("1:1048576").Hidden = True\nws.Rows(5).Hidden = False',
    "hide_rows_to_the_end": 'ws.Rows(2).RowHeight = 20\nws.Rows("3:1048576").Hidden = True',
    "hide_every_column": 'ws.Columns(3).ColumnWidth = 20\nws.Columns("A:XFD").Hidden = True',
    "font_rows": 'ws.Range("A1").Font.Size = 20\nws.Range("A2").Font.Name = "Segoe UI"\n'
                 'ws.Range("A3").Font.Name = "Arial"\nws.Range("A3").Font.Size = 26\nws.Range("A3").Font.Bold = True\n'
                 'ws.Rows(4).RowHeight = 12\nws.Range("A4").Font.Size = 20\nws.Range("A5").Value = 1\n'
                 'ws.Range("A5").Font.Size = 20\nws.Range("A5").ClearContents\nws.Range("A6").Font.Size = 20\n'
                 'ws.Rows(6).Hidden = True\nws.Range("B7").Font.Name = "Georgia"\nws.Range("B7").Font.Size = 40\n'
                 'ws.Range("A8").Font.Size = 14\nws.Range("B8").Font.Name = "Aptos"\nws.Range("B8").Font.Size = 30\n'
                 'ws.Range("A9").Font.Size = 12.375\nws.Range("A10").Font.Size = 8',
    # Heights the model cannot work out: a file Excel saved with them reads back
    # from the heights it recorded, which the model keeps while nothing changes.
    "recorded_rows": 'ws.Range("A1").Font.Size = 20\nws.Range("B1").Font.Name = "Arial"\nws.Range("B1").Font.Size = 26\n'
                     'ws.Range("A2").Font.Name = "Segoe Print"\nws.Range("A2").Font.Size = 20\n'
                     'ws.Range("A3").Font.Size = 20\nws.Range("A3").Font.Superscript = True\n'
                     'ws.Range("A4").Value = "x" & vbLf & "y"\nws.Range("A4").WrapText = True\n'
                     'ws.Range("A5").Value = "turned"\nws.Range("A5").Orientation = 90\n'
                     'ws.Range("A6").Font.Name = "Segoe Print"\nws.Range("A6").Font.Size = 8',
}

#: What is done to a workbook whose Normal font is Calibri 11, the default before Aptos.
CALIBRI_EDIT = ('ws.Range("A1").Value = 1\nws.Rows(3).RowHeight = 20\nws.Columns(2).ColumnWidth = 20\n'
                "ws.Rows(5).Hidden = True\nws.Columns(4).Hidden = True\nws.StandardWidth = 10")

#: (row, attributes) and (column, attributes) planted into a sheet before Excel opens it.
FOREIGN_ROWS = [(2, 'ht="20"'), (3, 'ht="12.34"'), (4, 'ht="0.1"'), (5, 'ht="409.5"'), (6, 'ht="500"'),
                (7, 'ht="20" customHeight="1"'), (8, 'ht="12.34" customHeight="1"'), (9, 'ht="0.1" customHeight="1"'),
                (10, 'ht="500" customHeight="1"'), (11, 'ht="-5" customHeight="1"'), (12, 'hidden="1"'),
                (13, 'ht="30" hidden="1" customHeight="1"'), (14, 'ht="0" customHeight="1"'), (15, 'ht="15"'),
                (16, 'ht="26.25"'), (17, 'customHeight="1"'), (18, 'ht="20.100000000000001" customHeight="1"'),
                (19, 'ht="7.5" customHeight="1"'), (20, 'ht="409.6" customHeight="1"'),
                (21, 'ht="409.7" customHeight="1"'), (22, 'ht="1000" customHeight="1"')]
FOREIGN_COLUMNS = [(2, 'width="15" customWidth="1"'), (3, 'width="12.345" customWidth="1"'), (4, 'width="15"'),
                   (5, 'width="0.5" customWidth="1"'), (6, 'width="300" customWidth="1"'),
                   (7, 'width="20" hidden="1" customWidth="1"'), (8, 'width="0" customWidth="1"'),
                   (9, 'width="9.140625" customWidth="1"'), (10, 'width="8.43" customWidth="1"'), (11, 'hidden="1"'),
                   (12, 'width="10.85546875" customWidth="1"'), (13, 'width="1" customWidth="1"'),
                   (14, 'width="1.7" customWidth="1"'), (15, 'width="0.1" customWidth="1"'),
                   (16, 'width="10.7874" customWidth="1"'), (17, 'width="-3" customWidth="1"')]
FOREIGN_READ_ROWS = 22
FOREIGN_READ_COLUMNS = 18


def sizes(xml: str) -> dict[str, object]:
    """The parts of a sheet that say how big its rows and columns are."""
    head = re.search(r"<sheetFormatPr\b[^>]*/>", xml)
    cols = re.search(r"<cols>.*?</cols>", xml, re.DOTALL)
    dimension = re.search(r"<dimension\b[^>]*/>", xml)
    return {"sheetFormatPr": head.group(0) if head else "", "cols": cols.group(0) if cols else "",
            "rows": re.findall(r"<row\b[^>]*>", xml), "dimension": dimension.group(0) if dimension else ""}


def sheet_parts(path: Path) -> list[str]:
    """Each worksheet's XML, in tab order."""
    with zipfile.ZipFile(path) as package:
        workbook = package.read("xl/workbook.xml").decode("utf-8")
        rels = package.read("xl/_rels/workbook.xml.rels").decode("utf-8")
        targets = dict(re.findall(r'Id="([^"]+)"[^>]*Target="([^"]+)"', rels))
        targets.update({rid: target for target, rid in re.findall(r'Target="([^"]+)"[^>]*Id="([^"]+)"', rels)})
        out = []
        for rid in re.findall(r'<sheet\b[^>]*r:id="([^"]+)"', workbook):
            target = targets[rid]
            part = target.lstrip("/") if target.startswith("/xl/") else f"xl/{target}"
            out.append(package.read(part).decode("utf-8"))
        return out


def run(excel: ExcelSession, code: str, proc: str) -> str:
    result = excel.run_vba(code, proc, timeout=600.0)
    if not result.ok:
        raise SystemExit(f"{proc}: {result.outcome}: {result.message} {result.error}")
    return str(result.value)


def main() -> None:
    from pyopenvba._templates import EMPTY_XLSM_BYTES

    FIXTURES.mkdir(parents=True, exist_ok=True)
    scratch = Path(tempfile.mkdtemp())
    authored = FIXTURES / "dimensions.xlsx"
    base = scratch / "base.xlsx"
    planted = FIXTURES / "foreign.xlsx"
    foreign_resaved = scratch / "foreign_resaved.xlsx"
    template = scratch / "template.xlsm"
    template.write_bytes(EMPTY_XLSM_BYTES)
    template_resaved = FIXTURES / "template_resaved.xlsx"
    calibri = FIXTURES / "calibri.xlsx"
    calibri_edited = scratch / "calibri_edited.xlsx"
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        standard = run(excel, "Public Function S() As String\nS = Workbooks.Add(xlWBATWorksheet).Worksheets(1)"
                              ".StandardHeight\nActiveWorkbook.Close False\nEnd Function\n", "S")
        assert standard == "15", f"needs a 96-DPI display (StandardHeight {standard})"
        # 1. Every case on its own sheet, saved, then read back from the file.
        lines = ["Public Function Build() As String", "Dim wb As Object, ws As Object", "Application.DisplayAlerts = False",
                 "Set wb = Workbooks.Add(xlWBATWorksheet)"]
        for index, (name, setup) in enumerate(CASES.items()):
            lines.append("Set ws = wb.Worksheets(1)" if index == 0 else
                         "Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))")
            lines.append(f'ws.Name = "{name}"')
            lines += setup.splitlines()
        lines += ["wb.Worksheets(1).Activate", f'wb.SaveAs Filename:="{authored}", FileFormat:=51', "wb.Close False",
                  "End Function"]
        run(excel, "\n".join(lines) + "\n", "Build")
        answers = run(excel, DESCRIBE + "Public Function ReadAll() As String\nDim wb As Object, ws As Object, out As String\n"
                                        f'Set wb = Workbooks.Open("{authored}")\n'
                                        'For Each ws In wb.Worksheets\nout = out & Describe(ws) & "|"\nNext ws\n'
                                        "wb.Close False\nReadAll = out\nEnd Function\n", "ReadAll")
        # 2. Sizes planted by hand, read and saved back by Excel.
        run(excel, "Public Function Base() As String\nDim wb As Object\nApplication.DisplayAlerts = False\n"
                   'Set wb = Workbooks.Add(xlWBATWorksheet)\nwb.Worksheets(1).Range("A1").Value = 1\n'
                   f'wb.SaveAs Filename:="{base}", FileFormat:=51\nwb.Close False\nEnd Function\n', "Base")
        with zipfile.ZipFile(base) as source:
            items = {name: source.read(name) for name in source.namelist()}
        xml = items["xl/worksheets/sheet1.xml"].decode("utf-8")
        cols = "<cols>" + "".join(f'<col min="{c}" max="{c}" {attrs}/>' for c, attrs in FOREIGN_COLUMNS) + "</cols>"
        rows = "".join(f'<row r="{r}" {attrs}/>' for r, attrs in FOREIGN_ROWS)
        xml = re.sub(r"(<sheetFormatPr[^>]*/>)", lambda m: m.group(1) + cols, xml, count=1)
        xml = re.sub(r"(<row r=\"1\"[^>]*>.*?</row>)", lambda m: m.group(1) + rows, xml, count=1, flags=re.DOTALL)
        items["xl/worksheets/sheet1.xml"] = xml.encode("utf-8")
        with zipfile.ZipFile(planted, "w", zipfile.ZIP_DEFLATED) as target:
            for name, data in items.items():
                target.writestr(name, data)
        foreign = run(excel, "Public Function Foreign() As String\nDim wb As Object, ws As Object, out As String\n"
                             "Dim r As Long, c As Long\nApplication.DisplayAlerts = False\n"
                             f'Set wb = Workbooks.Open("{planted}")\nSet ws = wb.Worksheets(1)\n'
                             f"For r = 1 To {FOREIGN_READ_ROWS}\n"
                             'out = out & ws.Rows(r).RowHeight & ":" & ws.Rows(r).Height & ":" & ws.Rows(r).Hidden & ":" '
                             '& ws.Rows(r).UseStandardHeight & ";"\nNext r\n'
                             f"For c = 1 To {FOREIGN_READ_COLUMNS}\n"
                             'out = out & ws.Columns(c).ColumnWidth & ":" & ws.Columns(c).Width & ":" & '
                             'ws.Columns(c).Hidden & ":" & ws.Columns(c).UseStandardWidth & ";"\nNext c\n'
                             'ws.Range("A1").Value = 2\n'
                             f'wb.SaveAs Filename:="{foreign_resaved}", FileFormat:=51\nwb.Close False\n'
                             "Foreign = out\nEnd Function\n", "Foreign")
        # 3. The new-workbook template, as Excel on this display writes it back.
        run(excel, "Public Function Template() As String\nDim wb As Object\nApplication.DisplayAlerts = False\n"
                   f'Set wb = Workbooks.Open("{template}")\nwb.Worksheets(1).Range("A1").Value = 1\n'
                   "wb.Worksheets(1).Rows(3).RowHeight = 20\n"
                   f'wb.SaveAs Filename:="{template_resaved}", FileFormat:=51\nwb.Close False\nEnd Function\n', "Template")
        # 4. A workbook whose Normal font is Calibri 11: saved empty, then edited and read back.
        run(excel, "Public Function Calibri() As String\nDim wb As Object\nApplication.DisplayAlerts = False\n"
                   'Set wb = Workbooks.Add(xlWBATWorksheet)\nwb.Styles("Normal").Font.Name = "Calibri"\n'
                   'wb.Styles("Normal").Font.Size = 11\n'
                   f'wb.SaveAs Filename:="{calibri}", FileFormat:=51\nwb.Close False\nEnd Function\n', "Calibri")
        calibri_answers = run(
            excel, DESCRIBE + "Public Function Edit() As String\nDim wb As Object, ws As Object\n"
                              f'Application.DisplayAlerts = False\nSet wb = Workbooks.Open("{calibri}")\n'
                              "Set ws = wb.Worksheets(1)\n" + CALIBRI_EDIT + "\n"
                              f'wb.SaveAs Filename:="{calibri_edited}", FileFormat:=51\nwb.Close False\n'
                              f'Set wb = Workbooks.Open("{calibri_edited}")\nEdit = Describe(wb.Worksheets(1))\n'
                              "wb.Close False\nEnd Function\n", "Edit")
    for fixture in (authored, planted, template_resaved, calibri):
        strip_save_path(fixture)
    authored_parts = sheet_parts(authored)
    described = [block for block in answers.split("|") if block]
    payload = {
        "rows": ROWS, "columns": COLUMNS, "describe": DESCRIBE,
        "cases": [{"name": name, "setup": setup, "answers": answer, "sizes": sizes(xml)}
                  for (name, setup), answer, xml in zip(CASES.items(), described, authored_parts, strict=True)],
        "foreign": {"rows": FOREIGN_ROWS, "columns": FOREIGN_COLUMNS, "read_rows": FOREIGN_READ_ROWS,
                    "read_columns": FOREIGN_READ_COLUMNS, "answers": foreign,
                    "resaved": sizes(sheet_parts(foreign_resaved)[0])},
        "template": sizes(sheet_parts(template_resaved)[0]),
        "calibri": {"edit": CALIBRI_EDIT, "answers": calibri_answers, "sizes": sizes(sheet_parts(calibri_edited)[0])},
    }
    (FIXTURES / "dimensions_answers.json").write_text(json.dumps(payload, indent=1) + "\n", encoding="utf-8")
    for case in payload["cases"]:
        print(case["name"], case["answers"][:120])
        print("   ", case["sizes"]["sheetFormatPr"], case["sizes"]["cols"][:300])
        print("   ", " ".join(case["sizes"]["rows"])[:400])
    print("foreign", foreign[:200])
    print("template", payload["template"])


if __name__ == "__main__":
    main()
