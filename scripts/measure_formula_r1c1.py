"""Measure native relative and absolute formula conversion."""
import json
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

PROBES = {
    "relative": '=RC[-2]+R[-1]C',
    "absolute": '=R1C1+R2C3',
    "mixed": '=R1C[-1]+R[-1]C1',
    "zero": '=R[0]C[0]',
    "range": '=SUM(R[-2]C[-2]:RC[-1])',
    "columns": '=SUM(C[-2]:C[-1])',
    "rows": '=SUM(R1:R2)',
    "quoted": '=IF(RC[-1]=1,"R1C1",RC[-2])',
    "sheet": "='Sheet1'!R1C1",
    "wrapped": '=R[-4]C[-4]',
    "constant": 'hello',
    "number": '12',
}


def main() -> None:
    records = []
    with ExcelSession(HarnessConfig(lock_wait_s=45.0)) as excel:
        excel.new_document()
        for name, formula in PROBES.items():
            body = ('Range("A1:F6").Clear\nOn Error Resume Next\n'
                    'Range("C3:D4").FormulaR1C1 = "' + formula.replace('"', '""') + '"\n'
                    'n = Err.Number\nOn Error GoTo 0\nReport = CStr(n)\n'
                    'For Each cell In Range("C3:D4")\n'
                    'Report = Report & "|" & cell.Formula & ":" & cell.FormulaR1C1 & ":" & TypeName(cell.FormulaR1C1)\nNext cell\n')
            code = 'Public Function Report() As String\nDim cell As Object, n As Long\n' + body + 'End Function'
            result = excel.run_vba(code, "Report", timeout=120.0)
            assert result.ok, result.message
            records.append(dict(name=name, body=body, reported=str(result.value)))
            print(name, result.value, flush=True)
    (Path(__file__).resolve().parents[1] / "tests/fixtures/formula_r1c1.json").write_text(
        json.dumps(records, indent=2) + "\n", encoding="utf-8")


if __name__ == "__main__":
    main()
