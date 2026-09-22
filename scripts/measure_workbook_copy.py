"""Measure workbook creation and sheet/range copy behavior."""
import json
from pathlib import Path
from pyvbaharness import ExcelSession, HarnessConfig

PROBES = {
    "range_values": 'src.Worksheets(1).Range("A1:B2").Copy dst.Worksheets(1).Range("C3")',
    "range_formula": 'src.Worksheets(1).Range("B1").Formula = "=A1+$A$1"\nsrc.Worksheets(1).Range("A1:B2").Copy dst.Worksheets(1).Range("C3")',
    "range_qualified": 'src.Worksheets(1).Range("B1").Formula = "=Sheet1!A1"\nsrc.Worksheets(1).Range("A1:B2").Copy dst.Worksheets(1).Range("C3")',
    "sheet_after": 'src.Worksheets(1).Copy After:=dst.Worksheets(1)',
    "sheet_before": 'src.Worksheets(1).Copy Before:=dst.Worksheets(1)',
    "sheet_new": 'src.Worksheets(1).Copy\nSet dst = ActiveWorkbook',
    "sheet_self_formula": 'src.Worksheets(1).Range("B1").Formula = "=Sheet1!A1"\nsrc.Worksheets(1).Copy After:=dst.Worksheets(1)',
    "sheet_names": 'src.Names.Add "Amount", "=Sheet1!$A$1"\nsrc.Worksheets(1).Names.Add "LocalAmount", "=Sheet1!$A$1"\nsrc.Worksheets(1).Range("B1").Formula = "=Amount+LocalAmount"\nsrc.Worksheets(1).Copy After:=dst.Worksheets(1)',
}

def main() -> None:
    records = []
    with ExcelSession(HarnessConfig(lock_wait_s=45.0)) as excel:
        for name, statement in PROBES.items():
            excel.new_document()
            body = ('Set src = ThisWorkbook\nSet dst = Workbooks.Add\n'
                    'src.Worksheets(1).Range("A1").Value = 10\nsrc.Worksheets(1).Range("B2").Value = 20\n'
                    + statement + '\nReport = CStr(dst.Worksheets.Count) & ":" & CStr(ActiveWorkbook Is dst) & "|"\n'
                    'For Each sh In dst.Worksheets\nReport = Report & sh.Name & ":"\n'
                    'For Each cell In sh.Range("A1:D4")\nIf cell.Formula <> "" Then Report = Report & cell.Address & "=" & cell.Formula & ";"\nNext cell\n'
                    'Report = Report & "|"\nNext sh\nFor Each nm In dst.Names\nReport = Report & nm.Name & "=" & nm.RefersTo & "|"\nNext nm\n')
            code = 'Function Report() As String\nDim src As Object, dst As Object, sh As Object, cell As Object, nm As Object\n' + body + 'End Function'
            result = excel.run_vba(code, "Report", timeout=120)
            assert result.ok, result.message
            records.append(dict(name=name, body=body, reported=str(result.value), source_name=excel.eval('ThisWorkbook.Name')))
            print(name, result.value, flush=True)
    (Path(__file__).resolve().parents[1] / "tests/fixtures/workbook_copy.json").write_text(json.dumps(records, indent=2) + "\n", encoding="utf-8")

if __name__ == "__main__":
    main()
