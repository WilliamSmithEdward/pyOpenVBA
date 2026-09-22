"""Measure native Excel Range.Find traversal and matching."""
import json
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parent.parent
PROBES = {
    "formula_boolean_case": 'Range("C3").Formula = "=TRUE()"\nSet hit = Range("C3").Find("TRUE", LookIn:=-4163, MatchCase:=True)',
    "saved_lookin": 'Range("C3").Formula = "=1+11"\nSet hit = Range("C3").Find("none", LookIn:=-4163)\nSet hit = Range("C3").Find(12)',
    "next_case": 'Set hit = Range("A1:C3").Find("alpha", MatchCase:=True)\nSet hit = Range("A1:C3").FindNext(hit)',
    "tilde": 'Range("C3").Value = "a~b"\nSet hit = Range("A1:C3").Find("~~")',
    "tilde_other": 'Range("C3").Value = "a~b"\nSet hit = Range("A1:C3").Find("~b")',
    "tilde_trailing": 'Range("C3").Value = "a~b"\nSet hit = Range("A1:C3").Find("~")',
    "tilde_trailing_after": 'Set hit = Range("A1:C3").Find("~", After:=Range("B1"))',
    "tilde_suffix": 'Set hit = Range("A1:C3").Find("alp~")',
    "tilde_trailing_whole": 'Set hit = Range("A1:C3").Find("~", LookAt:=1)',
    "error_text": 'Range("C3").Formula = "=1/0"\nSet hit = Range("C3").Find("#DIV/0!", LookIn:=-4163)',
    "all_star": 'Set hit = Range("A1:D4").Find("*", After:=Range("B1"))',
    "empty_single": 'Set hit = Range("D4").Find("", LookIn:=-4163)',
    "empty_formula_single": 'Range("C3").Formula = "="""""\nSet hit = Range("C3").Find("", LookIn:=-4163)',
    "empty_formula_text": 'Range("C3").Formula = "="""""\nSet hit = Range("C3").Find("", LookIn:=-4123)',
    "invalid_lookin": 'Set hit = Range("A1:C3").Find("alpha", LookIn:=0)',
    "invalid_lookat": 'Set hit = Range("A1:C3").Find("alpha", LookAt:=0)',
    "invalid_order": 'Set hit = Range("A1:C3").Find("alpha", SearchOrder:=0)',
    "invalid_direction": 'Set hit = Range("A1:C3").Find("alpha", SearchDirection:=0)',
    "next_no_after": 'Set hit = Range("A1:C3").Find("alpha")\nSet hit = Range("A1:C3").FindNext()',
    "previous_no_after": 'Set hit = Range("A1:C3").Find("alpha")\nSet hit = Range("A1:C3").FindPrevious()',
    "default": 'Set hit = Range("A1:C3").Find("alpha")',
    "whole": 'Set hit = Range("A1:C3").Find("alpha", LookAt:=1)',
    "case": 'Set hit = Range("A1:C3").Find("alpha", MatchCase:=True)',
    "after": 'Set hit = Range("A1:C3").Find("alpha", After:=Range("B1"))',
    "wrap": 'Set hit = Range("A1:C3").Find("Alpha", After:=Range("C3"), MatchCase:=True)',
    "columns": 'Set hit = Range("A1:C3").Find("alpha", SearchOrder:=2)',
    "previous": 'Set hit = Range("A1:C3").Find("alpha", SearchDirection:=2)',
    "missing": 'Set hit = Range("A1:C3").Find("absent")',
    "wildcard": 'Set hit = Range("A1:C3").Find("a?pha", LookAt:=1)',
    "star": 'Set hit = Range("A1:C3").Find("*bet*", LookAt:=1)',
    "literal_star": 'Range("C3").Value = "a*b"\nSet hit = Range("A1:C3").Find("~*")',
    "literal_question": 'Range("C3").Value = "a?b"\nSet hit = Range("A1:C3").Find("~?")',
    "number": 'Set hit = Range("A1:C3").Find(12, LookAt:=1)',
    "formula_text": 'Range("C3").Formula = "=1+11"\nSet hit = Range("A1:C3").Find("=1+11", LookIn:=-4123)',
    "formula_value": 'Range("C3").Formula = "=1+11"\nSet hit = Range("C3").Find(12, LookIn:=-4163)',
    "empty": 'Set hit = Range("A1:D4").Find("", LookIn:=-4163)',
    "empty_formula": 'Range("C3").Formula = "="""""\nSet hit = Range("A1:D4").Find("", LookIn:=-4163)',
    "boolean": 'Range("C3").Value = True\nSet hit = Range("A1:C3").Find(True, LookAt:=1)',
    "formatted_number": 'Range("C3").Value = 0.5\nRange("C3").NumberFormat = "0%"\nSet hit = Range("C3").Find("50%", LookIn:=-4163)',
    "raw_number": 'Range("C3").Value = 0.5\nRange("C3").NumberFormat = "0%"\nSet hit = Range("C3").Find("0.5", LookIn:=-4123)',
    "after_outside": 'Set hit = Range("A1:C3").Find("alpha", After:=Range("D4"))',
    "after_block": 'Set hit = Range("A1:C3").Find("alpha", After:=Range("A1:B1"))',
    "saved_whole": 'Set hit = Range("A1:C3").Find("none", LookAt:=1)\nSet hit = Range("A1:C3").Find("alp")',
    "saved_order": 'Set hit = Range("A1:C3").Find("none", SearchOrder:=2)\nSet hit = Range("A1:C3").Find("alpha")',
    "case_not_saved": 'Set hit = Range("A1:C3").Find("none", MatchCase:=True)\nSet hit = Range("A1:C3").Find("ALPHA")',
    "direction_not_saved": 'Set hit = Range("A1:C3").Find("none", SearchDirection:=2)\nSet hit = Range("A1:C3").Find("alpha")',
    "next": 'Set hit = Range("A1:C3").Find("alpha")\nSet hit = Range("A1:C3").FindNext(hit)',
    "previous_method": 'Set hit = Range("A1:C3").Find("alpha")\nSet hit = Range("A1:C3").FindPrevious(hit)',
}


def main() -> None:
    records = []
    with ExcelSession(HarnessConfig(lock_wait_s=45.0)) as excel:
        excel.new_document()
        for name, statement in PROBES.items():
            body = ('Range("A1:D4").Clear\nRange("A1").Value = "Alpha"\n'
                    'Range("B1").Value = "alphabet"\nRange("A2").Value = "alpha"\n'
                    'Range("C2").Value = 12\nRange("B3").Value = "ALPHA"\n'
                    'Set hit = Range("A1:C3").Find("__reset__", , -4123, 2, 1, 1, False, False, False)\n'
                    'On Error Resume Next\n' + statement + '\nn = Err.Number\nOn Error GoTo 0\n'
                    'Report = CStr(n) & "|"\nIf Not hit Is Nothing Then Report = Report & hit.Address\n')
            result = excel.run_vba('Public Function Report() As String\nDim hit As Object, n As Long\n'
                                   + body + 'End Function\n', "Report", timeout=120.0)
            assert result.ok, result.message
            records.append({"name": name, "body": body, "reported": str(result.value)})
            print(name, result.value, flush=True)
    (ROOT / "tests/fixtures/range_find.json").write_text(json.dumps(records, indent=2) + "\n", encoding="utf-8")


if __name__ == "__main__":
    main()
