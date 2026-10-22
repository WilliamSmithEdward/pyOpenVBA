"""Which formulas naming their own cell Excel takes for a circular reference.

A function that reads only where a reference is or how big it is, as ROWS
does, does not read the cells, and naming the formula's own cell that way
closes no loop. Each probe writes one formula to A5 of a new sheet whose
A1:A10 hold 1 to 10 but A5, with iteration off, and records the value and
the cell CircularReference names, or Nothing.

    python scripts/measure_self_reference.py

writes tests/fixtures/self_reference.json, which tests/test_excel_iteration.py
replays.
"""
from __future__ import annotations

from pathlib import Path

from vba_probe_fixture import record

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/self_reference.json'
FORMULAS = {
    'sum': '=SUM(A1:A10)',
    'rows': '=ROWS(A1:A10)',
    'columns': '=COLUMNS(A1:A10)',
    'row': '=ROW(A5)',
    'column': '=COLUMN(A5)',
    'row_range': '=ROW(A1:A10)',
    'areas': '=AREAS(A1:A10)',
    'isref': '=ISREF(A5)',
    'sheet': '=SHEET(A5)',
    'sheets': '=SHEETS(A5)',
    'cell_address': '=CELL("address",A5)',
    'cell_row': '=CELL("row",A5)',
    'cell_contents': '=CELL("contents",A5)',
    'cell_type': '=CELL("type",A5)',
    'formulatext': '=FORMULATEXT(A5)',
    'isformula': '=ISFORMULA(A5)',
    'index': '=INDEX(A1:A10,1)',
    'index_self': '=INDEX(A1:A10,5)',
    'offset': '=SUM(OFFSET(A5,1,0))',
    'offset_self': '=OFFSET(A4,1,0)',
    'indirect': '=INDIRECT("A4")',
    'indirect_self': '=INDIRECT("A5")',
    'counta': '=COUNTA(A1:A10)',
    'countblank': '=COUNTBLANK(A1:A10)',
    'isblank': '=ISBLANK(A5)',
    'n': '=N(A5)',
    'rows_in_sum': '=ROWS(A1:A10)+1',
    'rows_of_sum': '=ROWS(A1:A10)*SUM(A1)',
    'if_rows': '=IF(TRUE,ROWS(A1:A10),A5)',
    'if_false_branch': '=IF(TRUE,1,A5)',
    'choose_unused': '=CHOOSE(1,7,A5)',
    'iferror_unused': '=IFERROR(1,A5)',
    'let_unused': '=LET(x,A5,1)',
    'lookup': '=MATCH(3,A1:A10,0)',
    'vlookup_self': '=VLOOKUP(3,A1:A10,1,FALSE)',
    'columns_range_op': '=COLUMNS(A1:A10 A5)',
    'sumproduct_rows': '=SUMPRODUCT(ROW(A1:A10))',
}


def macro(formula: str) -> str:
    quoted = formula.replace('"', '""')
    return f'''Function Probe() As String
Dim ws As Worksheet, i As Long, s As String
Set ws = Worksheets.Add
For i = 1 To 10
    If i <> 5 Then ws.Cells(i, 1).Value = i
Next
ws.Range("A5").Formula = "{quoted}"
s = CStr(ws.Range("A5").Value) & "|"
If ws.CircularReference Is Nothing Then s = s & "Nothing" Else s = s & ws.CircularReference.Address
Application.DisplayAlerts = False
ws.Delete
Application.DisplayAlerts = True
Probe = s
End Function
'''


if __name__ == '__main__':
    record({name: macro(formula) for name, formula in FORMULAS.items()}, OUT)
