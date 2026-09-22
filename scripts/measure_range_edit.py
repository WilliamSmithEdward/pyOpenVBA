"""Measure whole-row/column edits and dependent formula references."""
import json
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

PROBES = {
    "insert_row": 'Rows("2:2").Insert',
    "insert_rows": 'Rows("2:3").Insert',
    "delete_row": 'Rows("2:2").Delete',
    "delete_rows": 'Rows("1:3").Delete',
    "insert_column": 'Columns("B:B").Insert',
    "insert_columns": 'Columns("B:C").Insert',
    "delete_column": 'Columns("B:B").Delete',
    "delete_columns": 'Columns("A:C").Delete',
    "insert_before_range": 'Rows("1:1").Insert',
    "insert_after_range": 'Rows("4:4").Insert',
    "delete_range_end": 'Rows("3:3").Delete',
    "delete_after_range": 'Rows("4:4").Delete',
}


def main() -> None:
    records = []
    with ExcelSession(HarnessConfig(lock_wait_s=45.0)) as excel:
        for name, statement in PROBES.items():
            excel.new_document()
            body = ('Worksheets.Add.Name = "Other"\nWorksheets("Sheet1").Activate\n'
                    'Range("A1").Value = 10\nRange("B2").Value = 20\nRange("C3").Value = 30\n'
                    'Range("E5").Formula = "=A1+$B$2+C$3+$C3"\n'
                    'Range("F5").Formula = "=SUM(A1:C3)"\n'
                    'Range("G5").Formula = "=SUM($A$1:$C$3)"\n'
                    'Range("H5").Formula = "=SUM(A:C)+SUM(1:3)"\n'
                    'Range("E6").Formula = "=IF(TRUE,""A1:B2"",Other!A1)"\n'
                    'Worksheets("Other").Range("A1").Formula = "=Sheet1!$B$2"\n'
                    'ThisWorkbook.Names.Add "Tracked", "=Sheet1!$A$1:$C$3"\n'
                    'On Error Resume Next\n' + statement + '\nn = Err.Number\nOn Error GoTo 0\n'
                    'Report = CStr(n)\nFor Each cell In Range("A1:J9")\n'
                    'If cell.Formula <> "" Then Report = Report & "|" & cell.Address & ":" & cell.Formula\n'
                    'Next cell\nReport = Report & "|Other:" & Worksheets("Other").Range("A1").Formula\n'
                    'Report = Report & "|Name:" & ThisWorkbook.Names("Tracked").RefersTo\n')
            result = excel.run_vba('Function Report() As String\nDim cell As Object, n As Long\n' + body + 'End Function', "Report", timeout=120.0)
            assert result.ok, result.message
            records.append(dict(name=name, body=body, reported=str(result.value)))
            print(name, result.value, flush=True)
    (Path(__file__).resolve().parents[1] / "tests/fixtures/range_edit.json").write_text(
        json.dumps(records, indent=2) + "\n", encoding="utf-8")


if __name__ == "__main__":
    main()
