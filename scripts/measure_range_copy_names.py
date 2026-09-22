"""Measure unattended Excel range-copy name conflict handling."""
import json
from pathlib import Path
from pyvbaharness import ExcelSession, HarnessConfig


def main() -> None:
    records = []
    with ExcelSession(HarnessConfig(lock_wait_s=45.0)) as excel:
        for kind in ("constant", "constant_conflict", "range", "range_conflict", "local", "unused", "different_sheet", "local_conflict"):
            excel.new_document()
            definition = "=7" if kind.startswith("constant") else "=Sheet1!$A$1"
            collection = 'src.Worksheets(1)' if kind.startswith("local") else 'src'
            setup = f'{collection}.Names.Add "Amount", "{definition}"\n'
            if kind.endswith("conflict"):
                setup += 'dst.Names.Add "Amount", "=99"\n'
            if kind == "different_sheet":
                setup += 'dst.Worksheets(1).Name = "Target"\n'
            body = ('Set src = ThisWorkbook\nSet dst = Workbooks.Add\n'
                    'Application.DisplayAlerts = False\n'
                    'src.Worksheets(1).Range("A1").Value = 7\n' + setup +
                    ('src.Worksheets(1).Range("B1").Formula = "=Amount+1"\n' if kind != "unused" else '') +
                    'src.Worksheets(1).Range("A1:B1").Copy dst.Worksheets(1).Range("C3")\n'
                    'Report = dst.Worksheets(1).Range("D3").Formula & "|" & CStr(dst.Worksheets(1).Range("D3").Value) & "|"\n'
                    'For Each nm In dst.Names\nReport = Report & nm.Name & "=" & nm.RefersTo & "|"\nNext nm\n')
            result = excel.run_vba('Function Report() As String\nDim src As Object, dst As Object, nm As Object\n' + body + 'End Function', 'Report', timeout=120)
            assert result.ok, result.message
            records.append(dict(name=kind, body=body, reported=str(result.value)))
            print(kind, result.value, flush=True)
    (Path(__file__).resolve().parents[1] / 'tests/fixtures/range_copy_names.json').write_text(json.dumps(records, indent=2) + '\n', encoding='utf-8')


if __name__ == '__main__':
    main()
