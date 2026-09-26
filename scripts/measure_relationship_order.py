"""The order Excel writes a part's relationships in, for every count from 1 to 160 and larger ones up to the most
hyperlinks a sheet holds.

Excel saves a workbook whose sheets hold 1, 2, ... 160 hyperlinks, then
200, 256, 300 and so on up to 65,530, one relationship each, numbered
rId1 to rIdn in the order the sheet lists them; the order each sheet's
.rels part names them in is recorded. The large sheets fill Office's
table to every size of 2**m buckets up to 16,384, with splits pending and
without, so their orders reach every bit of the hash a sheet can use.

    python scripts/measure_relationship_order.py

writes tests/fixtures/relationships/orders.json, which
tests/test_relationship_order.py checks src/pyopenvba/_relationships.py
against.
"""

from __future__ import annotations

import json
import re
import tempfile
import zipfile
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "tests" / "fixtures" / "relationships"
#: Every count to 160, then larger ones: each full level of buckets, counts between them, and 65,530, the most
#: hyperlinks Excel lets a sheet hold.
COUNTS = [*range(1, 161), 200, 256, 300, 500, 512, 600, 1000, 1024, 2000, 2048, 4096, 5000, 8192, 10000, 16384,
          32768, 65530]

MODULE = f"""Public Function Probe(target As String) As String
    Dim wb As Object, ws As Object, counts As Variant, k As Long, i As Long
    Application.ScreenUpdating = False
    counts = Array({", ".join(str(count) for count in COUNTS)})
    Set wb = Workbooks.Add(-4167)
    For k = 0 To UBound(counts)
        If k > 0 Then Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count)) Else Set ws = wb.Worksheets(1)
        For i = 1 To counts(k)
            ws.Hyperlinks.Add ws.Cells(i, 1), "http://example.com/" & k & "/" & i
        Next
    Next
    wb.SaveAs Filename:=target, FileFormat:=51
    wb.Close False
    Probe = "ok"
End Function
"""


def main() -> None:
    target = Path(tempfile.mkdtemp()) / "orders.xlsx"
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(MODULE, "Probe", args=(str(target),), timeout=1800.0)
    assert result.ok, f"{result.outcome}: {result.message}"
    orders: dict[int, list[int]] = {}
    with zipfile.ZipFile(target) as package:
        for name in package.namelist():
            if name.startswith("xl/worksheets/_rels/"):
                text = package.read(name).decode("utf-8")
                order = [int(one) for one in re.findall(r'<Relationship\b[^>]*?\bId="rId(\d+)"', text)]
                orders[len(order)] = order
    assert sorted(orders) == COUNTS, sorted(set(COUNTS) - set(orders))
    OUT.mkdir(parents=True, exist_ok=True)
    record = {"module": MODULE, "orders": {str(count): orders[count] for count in COUNTS}}
    (OUT / "orders.json").write_text(json.dumps(record, separators=(",", ":")) + "\n", encoding="utf-8")
    print(f"{len(orders)} orders written")


if __name__ == "__main__":
    main()
