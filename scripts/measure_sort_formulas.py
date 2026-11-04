"""What Range.Sort does to the references of a formula it moves, where Range.Copy would give #REF!.

Rows: A1:A5 hold 1 to 5 and a descending sort on them sends row r to row 6 - r. Columns: A1:E1 the same across. Each case puts a
formula in one cell of the sorted block, sorts, and records the formula where it lands and its value. Rival rules, written down
first: the moved reference is #REF! as in a copy; each end is clipped to the sheet's edge; each end wraps round the sheet.

    python scripts/measure_sort_formulas.py

writes tests/fixtures/sort_formulas.json, which tests/test_excel_sort_formulas.py replays.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/sort_formulas.json'
#: (cell the formula starts in, the formula, the cell it lands in); row r goes to 6 - r.
ROWS = (('C3', '=SUM(A1:A5)', 'C3'), ('C4', '=SUM(A1:A5)', 'C2'), ('C5', '=SUM(A1:A5)', 'C1'), ('C5', '=SUM(A1:A2)', 'C1'), ('C5', '=SUM(A3:A5)', 'C1'), ('C4', '=SUM(A2:A5)', 'C2'), ('C4', '=SUM(A3:A5)', 'C2'),
        ('C5', '=A1', 'C1'), ('C5', '=A3', 'C1'), ('C5', '=A5', 'C1'), ('C4', '=A1+A2', 'C2'), ('C4', '=A1', 'C2'), ('C5', '=$A1', 'C1'), ('C5', '=A$1', 'C1'), ('C5', '=$A$1', 'C1'), ('C5', '=SUM($A$1:$A$5)', 'C1'), ('C5', '=SUM($A$1:A5)', 'C1'), ('C5', '=SUM(A$1:A5)', 'C1'), ('C5', '=SUM(A1:A$5)', 'C1'),
        ('C1', '=SUM(A1:A5)', 'C5'), ('C1', '=A1', 'C5'), ('C1', '=A1048572', 'C5'), ('C1', '=A1048576', 'C5'), ('C1', '=SUM(A1048570:A1048576)', 'C5'), ('C1', '=SUM(A1048570:A1048572)', 'C5'), ('C1', '=SUM(A1:A1048576)', 'C5'), ('C2', '=A1048576', 'C4'),
        ('C5', '=H1', 'C1'), ('C5', '=H9', 'C1'), ('C5', '=SUM(H1:H9)', 'C1'), ('C5', '=SUM(A:A)', 'C1'), ('C5', '=SUM(5:5)', 'C1'), ('C5', '=SUM(1:5)', 'C1'), ('C5', '=SUM(1:1)', 'C1'), ('C1', '=SUM(1048576:1048576)', 'C5'),
        ('C5', '=One!A1', 'C1'), ('C5', '=One!A5', 'C1'), ('C5', '=SUM(One!A1:A5)', 'C1'), ('C5', '=Two!A1', 'C1'), ('C5', '=Two!A5', 'C1'), ('C5', '=SUM(Two!A1:A5)', 'C1'), ('C5', '=SUM(Two!A1:A2)', 'C1'), ('C5', '=Two!$A$1', 'C1'), ('C5', '=Nm', 'C1'), ('C5', '=ROW()', 'C1'), ('C5', '=OFFSET(A1,1,0)', 'C1'), ('C5', '=INDEX(A1:A5,2)', 'C1'), ('C5', '=SUM(A1:A2,A3:A5)', 'C1'), ('C5', '=SUM(A1:A2 A2:A3)', 'C1'), ('C5', '=IF(A1>0,SUM(A1:A5),A2)', 'C1'))
#: Across: column c goes to column 6 - c on a row-wise sort by row 1.
COLUMNS = (('E3', '=SUM(A1:E1)', 'A3'), ('E3', '=A1', 'A3'), ('E3', '=SUM(A1:B1)', 'A3'), ('E3', '=SUM(C1:E1)', 'A3'), ('A3', '=XFD1', 'E3'), ('A3', '=SUM(XFA1:XFD1)', 'E3'))
CELLS = 'Dim wb As Workbook, ws As Worksheet, i As Long\n    Application.DisplayAlerts = False\n    Set wb = Workbooks.Add(xlWBATWorksheet)\n    Set ws = wb.Worksheets(1)\n    ws.Name = "One"\n    wb.Worksheets.Add(After:=ws).Name = "Two"\n    wb.Names.Add "Nm", "=One!$A$1"\n    For i = 1 To 5\n        wb.Worksheets("Two").Cells(i, 1).Value = i * 100\n    Next\n    On Error GoTo Failed\n'


def source() -> str:
    lines: list[str] = []
    for label, cases, fill, sort in (
            ('Down', ROWS, '    For i = 1 To 5\n        ws.Cells(i, 1).Value = i\n    Next\n', 'ws.Range("A1:C5").Sort Key1:=ws.Range("A1"), Order1:=xlDescending, Header:=xlNo'),
            ('Across', COLUMNS, '    For i = 1 To 5\n        ws.Cells(1, i).Value = i\n    Next\n', 'ws.Range("A1:E3").Sort Key1:=ws.Range("A1"), Order1:=xlDescending, Header:=xlNo, Orientation:=xlSortRows')):
        for index, (start, formula, landing) in enumerate(cases):
            doubled = formula.replace('"', '""')
            lines += [f'Private Function {label}{index}() As String', '    ' + CELLS.rstrip('\n'), fill.rstrip('\n'), f'    ws.Range("{start}").Formula = "{doubled}"', f'    {sort}',
                      f'    {label}{index} = ws.Range("{landing}").Formula & "|" & ws.Range("{landing}").Text', '    GoTo Finish', 'Failed:', f'    {label}{index} = "Err " & Err.Number', '    Resume Finish', 'Finish:', '    wb.Close False', '    Application.DisplayAlerts = True', 'End Function', '']
    every = [f'{label}{index}' for label, cases in (('Down', ROWS), ('Across', COLUMNS)) for index in range(len(cases))]
    lines += ['Public Function Probe(ByVal first As Long, ByVal last As Long) As String', '    Dim i As Long, s As String', '    For i = first To last', '        Select Case i']
    lines += [f'            Case {index}: s = s & {name}() & "@@@"' for index, name in enumerate(every)]
    lines += ['        End Select', '    Next', '    Probe = s', 'End Function']
    return '\n'.join(lines) + '\n'


def main() -> None:
    cases = [['rows', *case] for case in ROWS] + [['columns', *case] for case in COLUMNS]
    answers: list[str] = []
    text = source()
    with ExcelSession(HarnessConfig(lock_wait_s=600.0)) as excel:
        excel.new_document()
        for first in range(0, len(cases), 15):
            done = excel.run_vba(text, 'Probe', (first, min(first + 14, len(cases) - 1)), timeout=200.0, line_numbers=False)
            assert done.ok, done
            answers += str(done.value).split('@@@')[:-1]
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'cases': [[*case, answer] for case, answer in zip(cases, answers, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for case in record['cases']:
        print(case[1], case[2], '->', case[3], '=>', case[4])


if __name__ == '__main__':
    main()
