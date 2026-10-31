"""Measure three small Range members: ID, which a macro sets and reads, AllowEdit on protected and unprotected
sheets, and ListHeaderRows, Excel's guess at how many header rows a list has.

Each case runs on a new workbook and answers as text.

    python scripts/measure_range_small_members.py

writes tests/fixtures/range_small_members.json, which tests/test_excel_range_small_members.py replays.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/range_small_members.json'
#: Rows of a list written from A1, each a VBA Array(...) body.
LISTS: list[list[str]] = [
    ['"Name", "Qty"', '"a", 1', '"b", 2'],
    ['"Name", "Qty"', '"a", "x"', '"b", "y"'],
    ['1, 2', '3, 4', '5, 6'],
    ['"Year", "Qty"', '2020, 1', '2021, 2'],
    ['2020, 2021', '1, 2', '3, 4'],
    ['"Name", "Qty"'],
    ['"Name", "Qty"', '"a", 1'],
    ['"Name", ""', '"a", 1', '"b", 2'],
    ['"", "Qty"', '"a", 1', '"b", 2'],
    ['"Name", "Qty"', '"", ""', '"b", 2'],
    ['"Name", "Qty"', '"Sub", "Unit"', '"a", 1', '"b", 2'],
    ['"name", "qty"', '"a", 1'],
    ['"Name", 5', '"a", 1', '"b", 2'],
    ['True, False', '1, 2'],
    ['"Name", "Qty"', '"a", #1/1/2020#', '"b", #1/2/2020#'],
    ['"Name", "Qty"', '"=1+1", 1'],
]
#: Each case: statements on a new workbook wb with sheet ws; the expression answered.
CASES: list[tuple[list[str], str]] = [
    ([], 'ws.Range("A1").ID'),
    (['ws.Range("A1").ID = "first"'], 'ws.Range("A1").ID'),
    (['ws.Range("A1").ID = "first"'], 'ws.Range("A1:B2").ID'),
    (['ws.Range("A1:B2").ID = "both"'], 'ws.Range("A1").ID & "|" & ws.Range("B2").ID & "|" & ws.Range("A1:B2").ID'),
    (['ws.Range("A1").ID = 5'], 'TypeName(ws.Range("A1").ID) & ":" & ws.Range("A1").ID'),
    (['ws.Range("A1").ID = "x"', 'ws.Range("A1").ClearContents'], 'ws.Range("A1").ID'),
    (['ws.Range("A1").ID = "x"', 'ws.Range("A1").Clear'], 'ws.Range("A1").ID'),
    (['ws.Range("A1").ID = "x"', 'ws.Range("A1").Copy ws.Range("C1")'], 'ws.Range("C1").ID'),
    (['ws.Range("A1").ID = "x"', 'ws.Rows(1).Insert'], 'ws.Range("A1").ID & "|" & ws.Range("A2").ID'),
    (['ws.Range("A1").ID = "x"', 'ws.Range("A2").ID = "x"'], 'ws.Range("A2").ID'),
    (['ws.Range("A1").ID = ""'], 'ws.Range("A1").ID'),
    (['ws.Range("A1").ID = "a b"'], 'ws.Range("A1").ID'),
    (['ws.Range("A1,C3").ID = "two"'], 'ws.Range("A1").ID & "|" & ws.Range("C3").ID'),
    ([], 'ws.Range("A1").AllowEdit'),
    (['ws.Protect'], 'ws.Range("A1").AllowEdit'),
    (['ws.Range("A1").Locked = False', 'ws.Protect'], 'ws.Range("A1").AllowEdit'),
    (['ws.Range("A1").Locked = False', 'ws.Protect'], 'ws.Range("A1:A2").AllowEdit'),
    (['ws.Range("A1:A2").Locked = False', 'ws.Protect'], 'ws.Range("A1:A2").AllowEdit'),
    (['ws.Protect UserInterfaceOnly:=True'], 'ws.Range("A1").AllowEdit'),
    (['ws.Protect DrawingObjects:=True, Contents:=False'], 'ws.Range("A1").AllowEdit'),
    (['ws.Protect', 'ws.Unprotect'], 'ws.Range("A1").AllowEdit'),
    (['ws.Range("A1").Locked = False', 'ws.Protect'], 'ws.Range("A1,C3").AllowEdit'),
    (['ws.Range("A1,C3").Locked = False', 'ws.Protect'], 'ws.Range("A1,C3").AllowEdit'),
    (['wb.Protect Structure:=True'], 'ws.Range("A1").AllowEdit'),
    (['ws.Protect Password:="p"'], 'ws.Range("A1").AllowEdit'),
    (['ws.Range("A1").Locked = Null', 'ws.Protect'], 'ws.Range("A1").AllowEdit'),
]
for rows in LISTS:
    statements = [f'ws.Range("A{at}").Resize(1, {len(rows[0].split(","))}).Value = Array({row})'
                  for at, row in enumerate(rows, start=1)]
    width = len(rows[0].split(','))
    CASES.append((statements, f'ws.Range("A1").Resize({len(rows)}, {width}).ListHeaderRows'))
CASES += [
    ([f'ws.Range("A{at}").Resize(1, 2).Value = Array({row})' for at, row in enumerate(LISTS[0], start=1)],
     'ws.Range("A2").ListHeaderRows'),
    ([f'ws.Range("A{at}").Resize(1, 2).Value = Array({row})' for at, row in enumerate(LISTS[0], start=1)],
     'ws.Range("A1").ListHeaderRows'),
    ([f'ws.Range("A{at}").Resize(1, 2).Value = Array({row})' for at, row in enumerate(LISTS[0], start=1)],
     'ws.Range("A1:A3").ListHeaderRows'),
    ([f'ws.Range("A{at}").Resize(1, 2).Value = Array({row})' for at, row in enumerate(LISTS[0], start=1)],
     'ws.Range("B1:B3").ListHeaderRows'),
    ([f'ws.Range("A{at}").Resize(1, 2).Value = Array({row})' for at, row in enumerate(LISTS[0], start=1)],
     'ws.Range("A2:B3").ListHeaderRows'),
    ([f'ws.Range("A{at}").Resize(1, 2).Value = Array({row})' for at, row in enumerate(LISTS[0], start=1)],
     'ws.Range("A1:B3,D1:E3").ListHeaderRows'),
    ([f'ws.Range("A{at}").Resize(1, 2).Value = Array({row})' for at, row in enumerate(LISTS[0], start=1)]
     + ['ws.Range("A1:B1").Font.Bold = True'], 'ws.Range("A1:B3").ListHeaderRows'),
    ([f'ws.Range("A{at}").Resize(1, 2).Value = Array({row})' for at, row in enumerate(LISTS[1], start=1)]
     + ['ws.Range("A1:B1").Font.Bold = True'], 'ws.Range("A1:B3").ListHeaderRows'),
    ([f'ws.Range("A{at}").Resize(1, 2).Value = Array({row})' for at, row in enumerate(LISTS[0], start=1)]
     + ['ws.ListObjects.Add xlSrcRange, ws.Range("A1:B3"), , xlYes'], 'ws.Range("A1:B3").ListHeaderRows'),
    ([], 'ws.Range("A1:B3").ListHeaderRows'),
    ([], 'ws.Range("A1").ListHeaderRows'),
]


def _probe() -> str:
    lines = ['Private Function Tried(ByVal which As Long) As String',
             '    Dim wb As Workbook, ws As Worksheet', '    Set wb = Workbooks.Add(xlWBATWorksheet)',
             '    Set ws = wb.Worksheets(1)', '    On Error GoTo Failed', '    Select Case which']
    for index, (statements, read) in enumerate(CASES):
        lines.append(f'    Case {index}')
        lines += [f'        {statement}' for statement in statements]
        lines.append(f'        Tried = CStr({read})')
    lines += ['    End Select', 'Done:', '    wb.Close False', '    Exit Function', 'Failed:',
              '    Tried = "Err " & Err.Number & " " & Err.Description', '    Resume Done', 'End Function', '',
              'Public Function Probe(ByVal first As Long, ByVal last As Long) As String',
              '    Dim s As String, i As Long', '    Application.DisplayAlerts = False', '    For i = first To last',
              '        s = s & "@@@" & Tried(i)', '    Next', '    Application.DisplayAlerts = True',
              '    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def main() -> None:
    source = _probe()
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(source, 'Probe', (0, len(CASES) - 1), timeout=300.0)
    assert result.ok and not result.dialogs, result
    answers = str(result.value).split('@@@')
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source,
              'cases': [[statements, read, answer] for (statements, read), answer in zip(CASES, answers, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for index, (statements, read, answer) in enumerate(record['cases']):
        print(index, (statements[-1] if statements else '')[:60], '|', read[:50], '->', answer)


if __name__ == '__main__':
    main()
