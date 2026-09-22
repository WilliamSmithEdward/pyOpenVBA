"""Measure reference-returning formulas used by Forms control names."""
import json
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parent.parent
PROBES = {
    "indirect_r1c1_list": ('=INDIRECT("Sheet1!R1C10:R3C10",FALSE)', False, ''),
    "indirect_r1c1_link": ('=INDIRECT("Sheet1!R2C8",FALSE)', True, ''),
    "indirect_r1c1_unqualified": ('=INDIRECT("R1C10:R3C10",FALSE)', False, ''),
    "indirect_r1c1_quoted": ('=INDIRECT("\'Sheet1\'!R1C10:R3C10",0)', False, ''),
    "indirect_r1c1_lowercase": ('=INDIRECT("sheet1!r1c10:r3c10",FALSE)', False, ''),
    "indirect_r1c1_zero": ('=INDIRECT("Sheet1!R0C10",FALSE)', False, ''),
    "indirect_r1c1_a1": ('=INDIRECT("Sheet1!J1:J3",FALSE)', False, ''),
    "indirect_r1c1_missing_sheet": ('=INDIRECT("Missing!R1C10:R3C10",FALSE)', False, ''),
    "choose_list": ('=CHOOSE(1,Choices,Sheet1!$J$2:$J$3)', False, ''),
    "choose_second": ('=CHOOSE(2,Choices,Sheet1!$J$2:$J$3)', False, ''),
    "choose_link": ('=CHOOSE(2,Sheet1!$H$1,Sheet1!$H$2)', True, ''),
    "choose_dynamic": ('=CHOOSE(Sheet1!$K$1,Sheet1!$J$1,Sheet1!$J$2:$J$3,Choices)', False, 'Range("K1").Value = 2'),
    "choose_invalid": ('=CHOOSE(0,Choices,Sheet1!$J$2:$J$3)', False, ''),
    "choose_fraction": ('=CHOOSE(1.9,Choices,Sheet1!$J$2:$J$3)', False, ''),
    "choose_lazy": ('=CHOOSE(1,Choices,1/0)', False, ''),
    "choose_scalar": ('=CHOOSE(2,Choices,42)', False, ''),
    "if_true": ('=IF(TRUE,Choices,Sheet1!$J$2:$J$3)', False, ''),
    "if_false": ('=IF(FALSE,Choices,Sheet1!$J$2:$J$3)', False, ''),
    "if_link": ('=IF(FALSE,Sheet1!$H$1,Sheet1!$H$2)', True, ''),
    "if_dynamic": ('=IF(Sheet1!$K$1=3,Choices,Sheet1!$J$2:$J$3)', False, 'Range("K1").Value = 2'),
    "if_lazy": ('=IF(TRUE,Choices,1/0)', False, ''),
    "if_scalar": ('=IF(FALSE,Choices,42)', False, ''),
    "if_omitted": ('=IF(FALSE,Choices)', False, ''),
    "if_error": ('=IF(1/0,Choices,Sheet1!$J$2:$J$3)', False, ''),
    "index_cell": ('=INDEX(Sheet1!$J$1:$J$3,2)', False, ''),
    "index_link": ('=INDEX(Sheet1!$H$1:$H$2,2)', True, ''),
    "index_column": ('=INDEX(Sheet1!$J$1:$K$3,0,1)', False, ''),
    "index_row": ('=INDEX(Sheet1!$J$1:$K$3,2,0)', False, ''),
    "index_all": ('=INDEX(Sheet1!$J$1:$K$3,0,0)', False, ''),
    "index_alias": ('=INDEX(Choices,0)', False, ''),
    "index_offset": ('=INDEX(OFFSET(Sheet1!$J$1,0,0,3,1),0,1)', False, ''),
    "index_dynamic": ('=INDEX(Sheet1!$H$1:$H$2,Sheet1!$K$1)', True, 'Range("K1").Value = 2\nsh.ControlFormat.Value = 1'),
    "index_row_vector": ('=INDEX(Sheet1!$H$2:$I$2,1)', True, ''),
    "index_two_dimensional": ('=INDEX(Sheet1!$H$1:$I$2,2)', True, ''),
    "index_outside": ('=INDEX(Sheet1!$J$1:$J$3,4)', False, ''),
    "index_negative": ('=INDEX(Sheet1!$J$1:$J$3,-1)', False, ''),
    "index_fraction": ('=INDEX(Sheet1!$H$1:$H$2,1.9)', True, ''),
    "index_area_one": ('=INDEX(Sheet1!$J$1:$J$3,0,1,1)', False, ''),
    "index_area_two": ('=INDEX(Sheet1!$J$1:$J$3,0,1,2)', False, ''),
    "offset_list": ('=OFFSET(Sheet1!$J$1,0,0,3,1)', False, ''),
    "offset_link": ('=OFFSET(Sheet1!$H$1,1,0)', True, ''),
    "offset_dynamic": ('=OFFSET(Sheet1!$J$1,0,0,Sheet1!$K$1,1)', False, 'Range("K1").Value = 2'),
    "offset_counta": ('=OFFSET(Sheet1!$J$1,0,0,COUNTA(Sheet1!$J$1:$J$3),1)', False, 'Range("J3").ClearContents'),
    "offset_alias": ('=OFFSET(Choices,1,0,2,1)', False, ''),
    "offset_nested": ('=OFFSET(OFFSET(Sheet1!$J$1,1,0),0,0,2,1)', False, ''),
    "indirect_list": ('=INDIRECT("Sheet1!$J$1:$J$3")', False, ''),
    "indirect_link": ('=INDIRECT("Sheet1!$H$2")', True, ''),
    "indirect_dynamic": ('=INDIRECT(Sheet1!$K$2)', False, 'Range("K2").Value = "Sheet1!$J$2:$J$3"'),
    "indirect_alias": ('=INDIRECT("Choices")', False, ''),
    "offset_zero": ('=OFFSET(Sheet1!$J$1,0,0,0,1)', False, ''),
    "offset_invalid": ('=OFFSET(Sheet1!$J$1,-1,0)', False, ''),
}


def main() -> None:
    records = []
    with ExcelSession(HarnessConfig(lock_wait_s=45.0)) as excel:
        excel.new_document()
        for name, (formula, linked, after) in PROBES.items():
            escaped = formula.replace('"', '""')
            field = "LinkedCell" if linked else "ListFillRange"
            body = ('Range("H1:K3").ClearContents\n'
                    'Range("J1").Value = "a"\nRange("J2").Value = "b"\nRange("J3").Value = "c"\n'
                    'Range("K1").Value = 3\nRange("K2").Value = "Sheet1!$J$1:$J$3"\n'
                    'ActiveWorkbook.Names.Add "Choices", "=Sheet1!$J$1:$J$3"\n'
                    f'ActiveWorkbook.Names.Add "Dynamic", "{escaped}"\n'
                    'Set sh = ActiveSheet.Shapes.AddFormControl(6, 0, 0, 90, 60)\n'
                    'sh.ControlFormat.ListFillRange = "$J$1:$J$3"\nOn Error Resume Next\n'
                    f'sh.ControlFormat.{field} = "Dynamic"\nsh.ControlFormat.Value = 3\n'
                    + after + '\nn = Err.Number\nOn Error GoTo 0\n'
                    'Report = CStr(n) & "|" & sh.ControlFormat.LinkedCell & "|" & sh.ControlFormat.ListFillRange & "|" & _\n'
                    'CStr(sh.ControlFormat.Value) & "|" & CStr(sh.ControlFormat.ListCount) & "|" & _\n'
                    'CStr(Range("H1").Value) & "|" & CStr(Range("H2").Value) & "|" & _\n'
                    'CStr(Range("I1").Value) & "|" & CStr(Range("I2").Value)\n'
                    'If sh.ControlFormat.ListCount > 0 Then Report = Report & "|" & sh.ControlFormat.List(1)\n')
            code = ('Public Function Report() As String\nDim sh As Object, n As Long\n'
                    'Do While ActiveSheet.Shapes.Count > 0\nActiveSheet.Shapes(1).Delete\nLoop\n'
                    + body + 'End Function\n')
            result = excel.run_vba(code, "Report", timeout=120.0)
            assert result.ok, result.message
            records.append({"name": name, "body": body, "reported": str(result.value)})
            print(name, result.value, flush=True)
    (ROOT / "tests/fixtures/shapes/control_formula_names.json").write_text(
        json.dumps(records, indent=2) + "\n", encoding="utf-8")


if __name__ == "__main__":
    main()
