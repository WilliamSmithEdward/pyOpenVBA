"""How Excel's dynamic arrays behave to a macro: Formula2, spilling, #SPILL!, A1# and the @ of a legacy formula.

Each case writes to a fresh sheet, 1, 2 and 3 in A1:A3, and reports each
cell it asks about: Formula, Formula2, FormulaR1C1, Formula2R1C1, Value,
Text, HasFormula, HasSpill, HasArray, and the addresses of SpillParent,
SpillingToRange and CurrentArray, a property Excel refuses as ! and its
error number. Expressions over whole ranges follow, each read as its type
and value; then a workbook of every kind of spill and #SPILL! is saved,
kept whole and in the parts that hold its dynamic arrays, and another of
one spill alone, for the metadata it needs.

    python scripts/measure_dynamic_arrays.py

writes tests/fixtures/dynamic_arrays.json and dynamic_arrays.xlsx, which
tests/test_excel_dynamic_arrays.py replays.
"""

from __future__ import annotations

import json
import tempfile
import zipfile
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "tests" / "fixtures" / "dynamic_arrays.json"
WORKBOOK = ROOT / "tests" / "fixtures" / "dynamic_arrays.xlsx"

HELPER = '''Private Function Prop(o As Object, name As String) As String
    Dim v As Variant
    On Error GoTo Bad
    If name = "SpillParent" Or name = "SpillingToRange" Or name = "CurrentArray" Then
        Prop = CallByName(o, name, VbGet).Address(False, False)
    Else
        v = CallByName(o, name, VbGet)
        If IsError(v) Then
            Prop = CStr(v)
        Else
            Prop = TypeName(v) & ":" & CStr(v)
        End If
    End If
    Exit Function
Bad:
    Prop = "!" & Err.Number
End Function

Private Function Look(ws As Object, addresses As String) As String
    Dim one As Variant, name As Variant, out As String
    For Each one In Split(addresses, ",")
        out = out & one & "{"
        For Each name In Array("Formula", "Formula2", "FormulaR1C1", "Formula2R1C1", "Value", "Text", _
                               "HasFormula", "HasSpill", "HasArray", "SpillParent", "SpillingToRange", "CurrentArray")
            out = out & name & "=" & Prop(ws.Range(one), CStr(name)) & ";"
        Next
        out = out & "}"
    Next
    Look = out
End Function
'''

#: Each case: the writes, then the cells reported.
CASES: dict[str, tuple[str, str]] = {
    "sequence": ('ws.Range("E1").Formula2 = "=SEQUENCE(3)"', "E1,E2,E3,E4"),
    "reference": ('ws.Range("E1").Formula2 = "=A1:A3"', "E1,E2,E3"),
    "single": ('ws.Range("E1").Formula2 = "=SUM(A1:A3)"', "E1,E2"),
    "arithmetic": ('ws.Range("E1").Formula2 = "=A1:A3*2"', "E1,E3"),
    "row": ('ws.Range("E1").Formula2 = "={1,2,3}"', "E1,F1,G1"),
    "square": ('ws.Range("E1").Formula2 = "=SEQUENCE(2,2)"', "E1,F1,E2,F2"),
    "if_array": ('ws.Range("E1").Formula2 = "=IF(A1:A3>1,""y"",""n"")"', "E1,E2,E3"),
    "filter_nothing": ('ws.Range("E1").Formula2 = "=FILTER(A1:A3,A1:A3>5)"', "E1,E2"),
    "unique_text": ('ws.Range("E1").Formula2 = "=UNIQUE({""a"";""b"";""a""})"', "E1,E2,E3"),
    "xlookup_cell": ('ws.Range("E1").Formula2 = "=XLOOKUP(2,A1:A3,A1:A3)"', "E1,E2"),
    "blocked": ('ws.Range("E2").Value = "x"\nws.Range("E1").Formula2 = "=SEQUENCE(3)"', "E1,E2,E3"),
    "unblocked": ('ws.Range("E2").Value = "x"\nws.Range("E1").Formula2 = "=SEQUENCE(3)"\n'
                  'ws.Range("E2").ClearContents', "E1,E2,E3"),
    "blocked_later": ('ws.Range("E1").Formula2 = "=SEQUENCE(3)"\nws.Range("E2").Value = 9', "E1,E2,E3"),
    "child_cleared": ('ws.Range("E1").Formula2 = "=SEQUENCE(3)"\nws.Range("E2").ClearContents', "E1,E2"),
    "source_edited": ('ws.Range("E1").Formula2 = "=A1:A3"\nws.Range("A2").Value = 20', "E1,E2"),
    "grows": ('ws.Range("E1").Formula2 = "=SEQUENCE(A1)"\nws.Range("A1").Value = 4', "E1,E4,E5"),
    "shrinks": ('ws.Range("E1").Formula2 = "=SEQUENCE(A3)"\nws.Range("A3").Value = 1', "E1,E2,E3"),
    "spill_reference": ('ws.Range("E1").Formula2 = "=SEQUENCE(3)"\nws.Range("G1").Formula2 = "=SUM(E1#)"', "G1"),
    "spill_reference_legacy": ('ws.Range("E1").Formula2 = "=SEQUENCE(3)"\nws.Range("G1").Formula = "=SUM(E1#)"',
                               "G1"),
    "spill_reference_none": ('ws.Range("G1").Formula2 = "=SUM(A1#)"', "G1"),
    "legacy_sequence": ('ws.Range("E1").Formula = "=SEQUENCE(3)"', "E1,E2"),
    "legacy_reference": ('ws.Range("E1").Formula = "=A1:A3"', "E1,E2"),
    "legacy_reference_row": ('ws.Range("E2").Formula = "=A1:A3"', "E2"),
    "legacy_sum": ('ws.Range("E1").Formula = "=SUM(A1:A3)"', "E1"),
    "legacy_index": ('ws.Range("E1").Formula = "=INDEX(A1:A3,0)"', "E1"),
    "at_written": ('ws.Range("E2").Formula2 = "=@A1:A3"', "E2"),
    "r1c1": ('ws.Range("E1").Formula2R1C1 = "=R1C1:R3C1"', "E1,E2"),
    "legacy_r1c1": ('ws.Range("E1").FormulaR1C1 = "=R1C1:R3C1"', "E1"),
    "formula_array": ('ws.Range("E1:E3").FormulaArray = "=SEQUENCE(3)"', "E1,E2"),
    "block_written": ('ws.Range("E1:E2").Formula2 = "=A1*2"', "E1,E2"),
    "text": ('ws.Range("E1").Formula2 = "={""a"";""b""}"', "E1,E2"),
    "anchor_cleared": ('ws.Range("E1").Formula2 = "=SEQUENCE(3)"\nws.Range("E1").ClearContents', "E1,E2"),
    "spill_cleared": ('ws.Range("E1").Formula2 = "=SEQUENCE(3)"\nws.Range("E1:E3").ClearContents', "E1,E2"),
    "child_written_formula": ('ws.Range("E1").Formula2 = "=SEQUENCE(3)"\nws.Range("E2").Formula = "=1"', "E1,E2"),
    "child_written_empty": ('ws.Range("E1").Formula2 = "=SEQUENCE(3)"\nws.Range("E2").Value = ""', "E1,E2,E3"),
    "merged": ('ws.Range("E2:F2").Merge\nws.Range("E1").Formula2 = "=SEQUENCE(3)"', "E1,E2"),
    "off_the_edge": ('ws.Range("E1048575").Formula2 = "=SEQUENCE(3)"', "E1048575,E1048576"),
    "copied": ('ws.Range("E1").Formula2 = "=SEQUENCE(3)"\nws.Range("E1").Copy ws.Range("G1")', "G1,G2,G3"),
    "moved_out": ('ws.Range("E1").Formula2 = "=SEQUENCE(3)"\nws.Range("E1").Cut ws.Range("G1")', "E1,E2,G1,G3"),
    "other_spill": ('ws.Range("E1").Formula2 = "=SEQUENCE(3)"\nws.Range("D2").Formula2 = "=SEQUENCE(1,3)"',
                    "E1,E2,D2"),
    "empty_string_spill": ('ws.Range("E1").Formula2 = "={""a"";""""}"', "E1,E2"),
    "spill_of_blank": ('ws.Range("E1").Formula2 = "=A5:A6"', "E1,E2"),
}

#: Each: the writes, then expressions over ranges, each read as ! and an error number or its type and value.
EXPRESSIONS: dict[str, tuple[str, list[str]]] = {
    "ranges": ('ws.Range("E1").Formula2 = "=SEQUENCE(3)"', [
        'TypeName(ws.Range("E1:E3").Value)', 'ws.Range("E1:E3").Value(2, 1)', 'ws.Range("E1:E3").HasFormula',
        'ws.Range("E2:E3").HasFormula', 'ws.UsedRange.Address', 'ws.Range("E1").End(-4121).Address',
        'ws.Range("E1").CurrentRegion.Address', 'ws.Cells.SpecialCells(-4123).Address',
        'ws.Cells.SpecialCells(2).Address', 'IsEmpty(ws.Range("E2").Value)', 'Application.WorksheetFunction.CountA('
        'ws.Range("E1:E3"))', 'ws.Range("E1").SpillingToRange.Rows.Count', 'ws.Range("E1:E3").HasSpill',
        'ws.Range("E2:E4").HasSpill', 'ws.Range("E1").Formula2 = ws.Range("E1").Formula',
        'ws.Evaluate("ROWS(E1#)")', 'ws.Range("E1#").Address', 'ws.Range("E4").End(-4162).Address']),
}

#: A workbook saved with dynamic arrays: what each cell holds, and the parts the file keeps for them.
FILE_WRITES = """ws.Range("E1").Formula2 = "=SEQUENCE(3)"
ws.Range("F1").Formula2 = "=A1:A3*2"
ws.Range("G1").Formula2 = "=SUM(E1#)"
ws.Range("H1").Formula = "=A1:A3"
ws.Range("I2").Value = "x"
ws.Range("I1").Formula2 = "=SEQUENCE(3)"
ws.Range("J1").Formula2 = "=SUM(A1:A3)"
ws.Range("K1").Formula2 = "={""a"",""b""}"
ws.Range("K3").Formula = "=SUM(E1#)"
ws.Range("M1").Formula2 = "=XLOOKUP(2,A1:A3,A1:A3)"
ws.Range("N1").Formula2 = "=FILTER(A1:A3,A1:A3>5)"
ws.Range("O1").Formula2 = "=SEQUENCE(2,2)"
ws.Range("Q2:R2").Merge
ws.Range("Q1").Formula2 = "=SEQUENCE(3)"
ws.Range("S1048575").Formula2 = "=SEQUENCE(3)"
ws.Range("T1").Formula2 = "=A1:A3/0"
ws.Range("U1").Formula2 = "=SEQUENCE(1)/0"
"""


def _expressions(ws_writes: str, expressions: list[str]) -> str:
    lines = ["Public Function Report() As String", "Dim ws As Object, out As String, v As Variant",
             "Set ws = ActiveWorkbook.Worksheets.Add",
             'ws.Range("A1:A3").Value = Application.Transpose(Array(1, 2, 3))', ws_writes, "On Error Resume Next"]
    for expression in expressions:
        lines += ["Err.Clear", "v = Empty", f"v = {expression}",
                  'If Err.Number <> 0 Then out = out & "!" & Err.Number Else out = out & TypeName(v) & ":" & CStr(v)',
                  'out = out & "|"']
    lines += ["Report = out", "End Function"]
    return "\n".join(lines) + "\n"


#: A workbook with one spill and no error, whose metadata has nothing but the dynamic-array flag.
PLAIN_WRITES = 'ws.Range("E1").Formula2 = "=SEQUENCE(3)"\n'


def _saved(excel: ExcelSession, folder: Path, writes: str = FILE_WRITES) -> dict[str, str]:
    """The parts of a workbook Excel saved with dynamic arrays in it."""
    path = folder / "spills.xlsx"
    code = ("Public Function Report() As String\nDim wb As Object, ws As Object\nApplication.DisplayAlerts = False\n"
            "Set wb = Workbooks.Add(xlWBATWorksheet)\nSet ws = wb.Worksheets(1)\n"
            'ws.Range("A1:A3").Value = Application.Transpose(Array(1, 2, 3))\n' + writes
            + f'wb.SaveAs Filename:="{path}", FileFormat:=51\nwb.Close False\nReport = "saved"\nEnd Function\n')
    result = excel.run_vba(code, "Report", timeout=120.0)
    assert result.ok, f"{result.outcome}: {result.message}"
    if writes == FILE_WRITES:
        # A workbook of its own, not the probe's, so the file carries no macro; the model opens it.
        WORKBOOK.write_bytes(path.read_bytes())
    parts: dict[str, str] = {}
    with zipfile.ZipFile(path) as package:
        for name in package.namelist():
            if name.endswith((".xml", ".rels")) and ("sheet1" in name or "metadata" in name or "Content_Types" in name
                                                    or "richData" in name or name == "xl/_rels/workbook.xml.rels"):
                parts[name] = package.read(name).decode("utf-8")
    return parts


def main() -> None:
    records = []
    expressions: dict[str, dict[str, str]] = {}
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        for name, (writes, cells) in CASES.items():
            body = ("Set ws = ActiveWorkbook.Worksheets.Add\n"
                    'ws.Range("A1:A3").Value = Application.Transpose(Array(1, 2, 3))\n'
                    + writes + f'\nReport = Look(ws, "{cells}")\n')
            code = HELPER + "\nPublic Function Report() As String\nDim ws As Object\n" + body + "End Function\n"
            result = excel.run_vba(code, "Report", timeout=120.0)
            assert result.ok, f"{name}: {result.outcome} {result.message}"
            records.append({"name": name, "body": body, "reported": str(result.value)})
            print(name, result.value, flush=True)
        for name, (writes, asked) in EXPRESSIONS.items():
            result = excel.run_vba(_expressions(writes, asked), "Report", timeout=120.0)
            assert result.ok, f"{name}: {result.outcome} {result.message}"
            answers = str(result.value).split("|")[:-1]
            expressions[name] = {"writes": writes, **dict(zip(asked, answers, strict=True))}
            print(name, answers, flush=True)
        with tempfile.TemporaryDirectory() as folder:
            parts = _saved(excel, Path(folder))
        with tempfile.TemporaryDirectory() as folder:
            plain = _saved(excel, Path(folder), PLAIN_WRITES)
    OUT.write_text(json.dumps({"helper": HELPER, "cases": records, "expressions": expressions,
                               "file": {"writes": FILE_WRITES, "parts": parts},
                               "plain_file": {"writes": PLAIN_WRITES, "parts": plain}}, indent=1) + "\n",
                   encoding="utf-8")


if __name__ == "__main__":
    main()
