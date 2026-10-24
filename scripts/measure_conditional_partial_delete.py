"""What FormatCondition.Delete does to a rule reached through a range that holds only part of it.

Each layout is a workbook of its own: the numbers r * 10 + c in A1:D10,
one to three rules, and one Delete, through a cell of a rule's range, a
block of it, all of it, more than it, one area of two, a cell in the
middle of a block, or a rule second in a cell's FormatConditions; on a
cell rule, an expression, a top 10, a colour scale and a data bar; called
at once or on a FormatCondition kept from before, and the collection's own
Delete beside them. Each workbook lists its rules after the Delete, with
their TypeName, range, priority, first formula and fill, and is saved, so
the file shows the rules and the dxfs Excel wrote for them.

    python scripts/measure_conditional_partial_delete.py

writes tests/fixtures/conditional_partial_delete/: each workbook as Excel
saved it, and conditional_partial_delete.json, which
tests/test_excel_conditional_partial_delete.py replays.
"""

from __future__ import annotations

import json
import re
import shutil
import sys
import zipfile
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "scripts"))
from fixture_workbook import strip_save_path  # noqa: E402

OUT = ROOT / "tests" / "fixtures" / "conditional_partial_delete"
RED, GREEN, BLUE = "RGB(255, 0, 0)", "RGB(0, 255, 0)", "RGB(0, 0, 255)"


def rule(area: str, colour: str, kind: str = 'xlCellValue, xlGreater, "5"') -> str:
    return f'Set f = ws.Range("{area}").FormatConditions.Add({kind})\nf.Interior.Color = {colour}\n'


#: Each layout: its name, the rules it makes, and the Delete.
LAYOUTS: list[tuple[str, str, str]] = [
    ("first_cell", rule("A1:A3", RED), 'ws.Range("A1").FormatConditions(1).Delete'),
    ("middle_cell", rule("A1:A3", RED), 'ws.Range("A2").FormatConditions(1).Delete'),
    ("last_cell", rule("A1:A3", RED), 'ws.Range("A3").FormatConditions(1).Delete'),
    ("two_cells", rule("A1:A3", RED), 'ws.Range("A1:A2").FormatConditions(1).Delete'),
    ("all_of_it", rule("A1:A3", RED), 'ws.Range("A1:A3").FormatConditions(1).Delete'),
    ("more_than_it", rule("A1:A3", RED), 'ws.Range("A1:B5").FormatConditions(1).Delete'),
    ("whole_sheet", rule("A1:A3", RED), "ws.Cells.FormatConditions(1).Delete"),
    ("other_area", rule("A1:A3,C1:C3", RED), 'ws.Range("C1:C3").FormatConditions(1).Delete'),
    ("part_of_area", rule("A1:A3,C1:C3", RED), 'ws.Range("C1").FormatConditions(1).Delete'),
    ("hole", rule("A1:C3", RED), 'ws.Range("B2").FormatConditions(1).Delete'),
    ("corner", rule("A1:C3", RED), 'ws.Range("A1").FormatConditions(1).Delete'),
    ("second_rule", rule("A1:A3", RED) + rule("A1:A5", BLUE), 'ws.Range("A1").FormatConditions(2).Delete'),
    ("first_of_two", rule("A1:A3", RED) + rule("A1:A5", BLUE), 'ws.Range("A1").FormatConditions(1).Delete'),
    ("priorities", rule("A1:A3", RED) + rule("B1:B3", GREEN) + rule("C1:C3", BLUE),
     'ws.Range("B1").FormatConditions(1).Delete'),
    ("expression", rule("A1:A3", RED, 'xlExpression, , "=A1>5"'), 'ws.Range("A1").FormatConditions(1).Delete'),
    ("top10", 'Set f = ws.Range("A1:A10").FormatConditions.AddTop10\nf.Interior.Color = RGB(255, 0, 0)\n',
     'ws.Range("A1").FormatConditions(1).Delete'),
    ("colour_scale", 'Set f = ws.Range("A1:A10").FormatConditions.AddColorScale(3)\n',
     'ws.Range("A1").FormatConditions(1).Delete'),
    ("data_bar", 'Set f = ws.Range("A1:A10").FormatConditions.AddDatabar\n',
     'ws.Range("A1").FormatConditions(1).Delete'),
    ("kept", rule("A1:A3", RED), 'Set g = ws.Range("A1").FormatConditions(1)\nws.Range("D9").Value = 1\ng.Delete'),
    ("kept_whole", rule("A1:A3", RED), 'Set g = ws.Range("A1:A3").FormatConditions(1)\ng.Delete'),
    ("collection", rule("A1:A3", RED), 'ws.Range("A1").FormatConditions.Delete'),
    ("collection_priorities", rule("A1:A3", RED) + rule("B1:B3", GREEN) + rule("C1:C3", BLUE),
     'ws.Range("B1").FormatConditions.Delete'),
]

BUILD = '''Private Function Listed(ws As Object) As String
    Dim f As Object, out As String, formula As String, colour As String
    For Each f In ws.Cells.FormatConditions
        formula = ""
        colour = ""
        On Error Resume Next
        formula = f.Formula1
        colour = CStr(f.Interior.Color)
        On Error GoTo 0
        out = out & TypeName(f) & " " & f.AppliesTo.Address(False, False) & "@" & f.Priority & "=" & formula & _
            "/" & colour & ";"
    Next
    Listed = out
End Function

Public Function Build() As String
Dim wb As Object, ws As Object, f As Object, g As Object, out As String, r As Long, c As Long
Application.DisplayAlerts = False
{layouts}
Build = out
End Function
'''

LAYOUT = '''Set wb = Workbooks.Add(xlWBATWorksheet)
Set ws = wb.Worksheets(1)
For r = 1 To 10
    For c = 1 To 4
        ws.Cells(r, c).Value = r * 10 + c
    Next
Next
{rules}out = out & "{name}^" & Listed(ws) & "~"
On Error Resume Next
{delete}
If Err.Number <> 0 Then out = out & "E" & Err.Number & "~"
On Error GoTo 0
out = out & Listed(ws) & "|"
wb.SaveAs "{folder}\\{name}.xlsx", 51
wb.Close False
'''


def holdings(path: Path) -> dict[str, object]:
    """The sheet's conditional formats as the file keeps them, and the dxfs."""
    with zipfile.ZipFile(path) as package:
        sheet = package.read("xl/worksheets/sheet1.xml").decode("utf-8")
        styles = package.read("xl/styles.xml").decode("utf-8")
    groups = re.findall(r"<conditionalFormatting\b.*?</conditionalFormatting>", sheet, re.DOTALL)
    extension = re.findall(r"<x14:conditionalFormatting\b.*?</x14:conditionalFormatting>", sheet, re.DOTALL)
    dxfs = re.search(r"<dxfs\b[^>]*?(?:/>|>(.*?)</dxfs>)", styles, re.DOTALL)
    listed = re.findall(r"<dxf>.*?</dxf>|<dxf/>", dxfs.group(1) or "", re.DOTALL) if dxfs else []
    return {"rules": groups, "extension": extension, "dxfs": listed}


def main() -> None:
    if OUT.exists():
        shutil.rmtree(OUT)
    OUT.mkdir(parents=True)
    body = "".join(LAYOUT.replace("{name}", name).replace("{rules}", rules).replace("{delete}", delete)
                   for name, rules, delete in LAYOUTS)
    code = BUILD.replace("{layouts}", body).replace("{folder}", str(OUT))
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(code, "Build", timeout=600.0)
        assert result.ok, f"{result.outcome}: {result.message} {result.error}"
    shown = dict(one.split("^", 1) for one in str(result.value).split("|") if "^" in one)
    layouts: list[dict[str, object]] = []
    for name, rules, delete in LAYOUTS:
        strip_save_path(OUT / f"{name}.xlsx")
        before, after = shown[name].split("~", 1)
        layouts.append({"name": name, "rules": rules, "delete": delete, "before": before, "after": after,
                        "file": holdings(OUT / f"{name}.xlsx")})
    record = {"build": BUILD, "layout": LAYOUT, "layouts": layouts}
    (OUT / "conditional_partial_delete.json").write_text(json.dumps(record, indent=1, ensure_ascii=False) + "\n",
                                                        encoding="utf-8")
    for layout in layouts:
        print("==", layout["name"], "|", layout["delete"].splitlines()[-1])
        print("   before", layout["before"])
        print("   after ", layout["after"])
        print("   file  ", [re.sub(r"\s+", " ", group)[:160] for group in layout["file"]["rules"]],
              len(layout["file"]["dxfs"]), "dxfs")


if __name__ == "__main__":
    main()
