"""What Excel names each workbook a macro makes: Workbooks.Add with no template and with xlWBATWorksheet, in turn,
some closed between them, and a sheet copied into a new workbook.

The macro runs in the harness's own workbook, Book1, which Excel made first. It records how many workbooks are open
and the active one's name, then each new workbook's name in the order made.

    python scripts/measure_workbook_names.py

writes tests/fixtures/workbook_names.json, which tests/test_excel_workbook_names.py replays.
"""

from __future__ import annotations

import json
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parent.parent / "tests" / "fixtures" / "workbook_names.json"

CODE = '''
Public Function Names() As String
    Dim out As String, a As Object, b As Object, c As Object, d As Object, e As Object, f As Object
    out = "start:" & Workbooks.Count & ":" & ActiveWorkbook.Name & ";"
    Set a = Workbooks.Add
    Set b = Workbooks.Add(xlWBATWorksheet)
    Set c = Workbooks.Add
    Set d = Workbooks.Add(xlWBATWorksheet)
    out = out & a.Name & ";" & b.Name & ";" & c.Name & ";" & d.Name & ";"
    b.Close False
    Set e = Workbooks.Add(xlWBATWorksheet)
    a.Close False
    Set f = Workbooks.Add
    out = out & e.Name & ";" & f.Name & ";"
    c.Close False: d.Close False: e.Close False: f.Close False
    Set a = Workbooks.Add(xlWBATWorksheet)
    a.Worksheets(1).Copy
    out = out & "copied from " & a.Name & ":" & ActiveWorkbook.Name & ";"
    ActiveWorkbook.Close False
    Set b = Workbooks.Add
    b.Worksheets(1).Copy
    out = out & "copied from " & b.Name & ":" & ActiveWorkbook.Name & ";"
    ActiveWorkbook.Close False
    a.Close False: b.Close False
    Names = out
End Function
'''


def main() -> None:
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(CODE, "Names", timeout=120.0)
        assert result.ok, f"{result.outcome}: {result.message} {result.error}"
    OUT.write_text(json.dumps({"code": CODE, "names": str(result.value)}, indent=1) + "\n", encoding="utf-8")
    print(result.value)


if __name__ == "__main__":
    main()
