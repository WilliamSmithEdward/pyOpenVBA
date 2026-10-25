"""A pivot table's standard deviations and variances, as Excel works them out and saves them, to the last bit.

A macro builds one pivot table over groups of values chosen to set the ways
of working out a variance apart: large values a little apart, where the sum
of squares less the square of the sum loses most of its digits, decimals,
tiny values, values all the same, and a group of whole numbers. The table
shows each group's StdDev, StdDevP, Var and VarP and the grand total of all
of them, and is saved; the saved cells hold every value to the bit.

    python scripts/measure_pivot_variance.py

writes tests/fixtures/pivot_variance/: the workbook as Excel saved it and
pivot_variance.json, which tests/test_excel_pivot_variance.py replays.
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

OUT = ROOT / "tests" / "fixtures" / "pivot_variance"

#: Each group's values, in the order the source lists them.
GROUPS: dict[str, list[str]] = {
    "Close": ["100000001", "100000002", "100000003"],
    "Closer": ["1000000001", "1000000004", "1000000002", "1000000007"],
    "Decimals": ["0.1", "0.2", "0.3", "0.7"],
    "Tiny": ["1E-100", "3E-100", "2E-100"],
    "Same": ["1E-100", "1E-100", "1E-100"],
    "Whole": ["2", "5", "8", "5", "6"],
    "Mixed": ["1E+15", "1", "-1E+15", "3"],
}

BUILD = '''
Public Function Build() As String
Dim wb As Object, data As Object, cache As Object, pt As Object, r As Long, rows As Variant, parts As Variant
Application.DisplayAlerts = False
Set wb = Workbooks.Add(xlWBATWorksheet)
Set data = wb.Worksheets(1)
data.Name = "data"
data.Range("A1:B1").Value = Array("Group", "Value")
rows = Array({rows})
For r = 0 To UBound(rows)
    parts = Split(rows(r), "|")
    data.Cells(r + 2, 1).Value = parts(0)
    data.Cells(r + 2, 2).Value = CDbl(parts(1))
Next
Set cache = wb.PivotCaches.Create(xlDatabase, "data!R1C1:R" & (UBound(rows) + 2) & "C2")
Set pt = cache.CreatePivotTable(data.Range("E3"), "Spread")
pt.PivotFields("Group").Orientation = xlRowField
pt.AddDataField pt.PivotFields("Value"), "Sample deviation", xlStDev
pt.AddDataField pt.PivotFields("Value"), "Whole deviation", xlStDevP
pt.AddDataField pt.PivotFields("Value"), "Sample variance", xlVar
pt.AddDataField pt.PivotFields("Value"), "Whole variance", xlVarP
wb.SaveAs "{folder}\\pivot_variance.xlsx", 51
wb.Close False
Build = "done"
End Function
'''


def main() -> None:
    if OUT.exists():
        shutil.rmtree(OUT)
    OUT.mkdir(parents=True)
    rows = ", ".join(f'"{group}|{value}"' for group, values in GROUPS.items() for value in values)
    code = BUILD.replace("{rows}", rows).replace("{folder}", str(OUT))
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(code, "Build", timeout=300.0)
        assert result.ok, f"{result.outcome}: {result.message} {result.error}"
    strip_save_path(OUT / "pivot_variance.xlsx")
    (OUT / "pivot_variance.json").write_text(json.dumps({"build": BUILD, "groups": GROUPS}, indent=1) + "\n",
                                             encoding="utf-8")
    print("saved", OUT / "pivot_variance.xlsx")


if __name__ == "__main__":
    main()
