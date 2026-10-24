"""Which dxfs Excel's save keeps once nothing names them: a deleted rule's, a recoloured one's, a deleted sheet's.

tests/fixtures/dxf_table/ shows a save leaving out a dxf a file holds that
nothing names, and the dxf a rule had before a macro gave it a new format.
A rule deleted whole leaves its dxf in the save all the same
(tests/fixtures/conditional_partial_delete/). Each layout here is a
workbook of its own: rules made, then one taken away by
FormatCondition.Delete, by the collection's Delete over its range or over
every cell, by ClearFormats or Clear, by a paste over it, by deleting its
rows or its sheet, after a new colour, before a new rule, or in a
workbook saved and opened again; then saved, twice for one of them. Each file's dxfs and every name of one
are kept.

    python scripts/measure_dxf_lifetime.py

writes tests/fixtures/dxf_lifetime/: each workbook as Excel saved it, and
dxf_lifetime.json, which tests/test_excel_dxf_lifetime.py replays.
"""

from __future__ import annotations

import json
import shutil
import sys
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "scripts"))
from fixture_workbook import strip_save_path  # noqa: E402
from measure_dxf_table import holdings  # noqa: E402

OUT = ROOT / "tests" / "fixtures" / "dxf_lifetime"


def rule(sheet: str, area: str, colour: str) -> str:
    return (f'Set f = wb.Worksheets("{sheet}").Range("{area}").FormatConditions.Add(xlCellValue, xlGreater, "5")\n'
            f"f.Interior.Color = {colour}\n")


RED, GREEN, BLUE = "RGB(255, 0, 0)", "RGB(0, 255, 0)", "RGB(0, 0, 255)"
TWO = rule("one", "A1:A3", RED) + rule("one", "B1:B3", GREEN)
#: Each layout: its name, what it does once the workbook has its sheets, and the files it saves, the last one the
#: layout's own.
LAYOUTS: list[tuple[str, str, list[str]]] = [
    ("deleted", TWO + 'wb.Worksheets("one").Range("A1:A3").FormatConditions(1).Delete\n', []),
    ("recoloured_then_deleted", TWO + 'wb.Worksheets("one").Range("A1").FormatConditions(1).Interior.Color = '
     f'{BLUE}\nwb.Worksheets("one").Range("A1:A3").FormatConditions(1).Delete\n', []),
    ("collection", TWO + 'wb.Worksheets("one").Range("A1:A3").FormatConditions.Delete\n', []),
    ("every_cell", TWO + 'wb.Worksheets("one").Cells.FormatConditions.Delete\n', []),
    ("clear_formats", TWO + 'wb.Worksheets("one").Range("A1:A3").ClearFormats\n', []),
    ("cleared", TWO + 'wb.Worksheets("one").Range("A1:A3").Clear\n', []),
    ("pasted_over", TWO + 'wb.Worksheets("one").Range("D1:D3").Copy wb.Worksheets("one").Range("A1:A3")\n', []),
    ("rows_deleted", TWO + 'wb.Worksheets("one").Rows("1:3").Delete\n', []),
    ("sheet_deleted", TWO + rule("two", "A1:A3", BLUE) + 'wb.Worksheets("two").Delete\n', []),
    ("deleted_then_added", TWO + 'wb.Worksheets("one").Range("A1:A3").FormatConditions(1).Delete\n'
     + rule("one", "C1:C3", BLUE), []),
    ("saved_twice", TWO + 'wb.Worksheets("one").Range("A1:A3").FormatConditions(1).Delete\n', ["saved_twice_first"]),
    ("reopened", TWO + 'wb.Worksheets("one").Range("A1:A3").FormatConditions(1).Delete\n'
     'wb.SaveAs "{folder}\\reopened_first.xlsx", 51\nwb.Close False\n'
     'Set wb = Workbooks.Open("{folder}\\reopened_first.xlsx")\n', []),
    ("read_then_deleted", TWO + 'wb.SaveAs "{folder}\\read_then_deleted_first.xlsx", 51\nwb.Close False\n'
     'Set wb = Workbooks.Open("{folder}\\read_then_deleted_first.xlsx")\n'
     'wb.Worksheets("one").Range("A1:A3").FormatConditions(1).Delete\n', []),
]

BUILD = '''Public Function Build() As String
Dim wb As Object, f As Object, r As Long
Application.DisplayAlerts = False
{layouts}
Build = "ok"
End Function
'''

LAYOUT = '''Set wb = Workbooks.Add(xlWBATWorksheet)
wb.Worksheets(1).Name = "one"
wb.Worksheets.Add(After:=wb.Worksheets(1)).Name = "two"
wb.Worksheets("one").Activate
For r = 1 To 6
    wb.Worksheets("one").Cells(r, 1).Value = 9
    wb.Worksheets("one").Cells(r, 2).Value = 9
    wb.Worksheets("one").Cells(r, 3).Value = 9
    wb.Worksheets("two").Cells(r, 1).Value = 9
Next
{steps}{saves}wb.SaveAs "{folder}\\{name}.xlsx", 51
wb.Close False
'''


def main() -> None:
    if OUT.exists():
        shutil.rmtree(OUT)
    OUT.mkdir(parents=True)
    body = "".join(LAYOUT.replace("{steps}", steps).replace("{name}", name).replace(
        "{saves}", "".join(f'wb.SaveAs "{{folder}}\\{save}.xlsx", 51\n' for save in saves))
        for name, steps, saves in LAYOUTS)
    code = BUILD.replace("{layouts}", body).replace("{folder}", str(OUT))
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(code, "Build", timeout=600.0)
        assert result.ok and result.value == "ok", f"{result.outcome}: {result.message} {result.error}"
    files = sorted(path.stem for path in OUT.glob("*.xlsx"))
    for name in files:
        strip_save_path(OUT / f"{name}.xlsx")
    record = {"build": BUILD, "layout": LAYOUT,
              "layouts": [{"name": name, "steps": steps, "saves": saves} for name, steps, saves in LAYOUTS],
              "files": {name: holdings(OUT / f"{name}.xlsx") for name in files}}
    (OUT / "dxf_lifetime.json").write_text(json.dumps(record, indent=1, ensure_ascii=False) + "\n", encoding="utf-8")
    for name in files:
        held = record["files"][name]
        print("==", name, [one[-44:-27] for one in held["dxfs"]])
        print("  ", held["names"])


if __name__ == "__main__":
    main()
