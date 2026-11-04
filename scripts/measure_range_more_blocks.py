"""The parts of the VERIFIED standard scripts/measure_range_more.py left out: the local formula and format properties set
on a block, and each of its properties read over cells that differ. Nothing here reads a size, so unlike that script
this one runs at any display DPI.

    python scripts/measure_range_more_blocks.py

writes tests/fixtures/range_more_blocks.json, which tests/test_excel_range_more_blocks.py replays.
"""
from __future__ import annotations

import json
from pathlib import Path

import measure_range_more as more
from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/range_more_blocks.json'
C2C3 = 'ws.Range("C2:C3")'
#: Set on a block, as tests/fixtures/range_more.json has them on a cell.
BLOCKS = [
    ('FormulaLocal', C2C3, more.FORMULAS, ['FormulaLocal', 'Formula', 'HasFormula']),
    ('FormulaR1C1Local', C2C3, more.FORMULAS, ['FormulaR1C1Local', 'Formula', 'HasFormula']),
    ('NumberFormatLocal', C2C3, more.FORMATS, ['NumberFormatLocal', 'NumberFormat']),
]
#: A read over cells that differ: (member, what makes them differ, the range).
DIFFERING = [
    ('Formula2', 'ws.Range("C2").Formula = "=1+1": ws.Range("C3").Value = 5', C2C3),
    ('FormulaLocal', 'ws.Range("C2").Formula = "=1+1": ws.Range("C3").Value = 5', C2C3),
    ('FormulaR1C1', 'ws.Range("C2").Formula = "=1+1": ws.Range("C3").Value = 5', C2C3),
    ('FormulaR1C1Local', 'ws.Range("C2").Formula = "=1+1": ws.Range("C3").Value = 5', C2C3),
    ('FormulaArray', 'ws.Range("C2").FormulaArray = "=1+1": ws.Range("C3").Value = 5', C2C3),
    ('NumberFormatLocal', 'ws.Range("C2").NumberFormat = "0.00"', C2C3),
    ('MergeCells', 'ws.Range("B2:C2").Merge', 'ws.Range("B2:C3")'),
    ('Name', 'ws.Range("C2").Name = "Differ"', C2C3),
    ('ID', 'ws.Range("C2").ID = "one"', C2C3),
    ('Hidden', 'ws.Rows(2).Hidden = True', 'ws.Rows("2:3")'),
    ('Hidden', 'ws.Columns(2).Hidden = True', 'ws.Columns("B:C")'),
    ('OutlineLevel', 'ws.Rows(2).OutlineLevel = 3', 'ws.Rows("2:3")'),
    ('OutlineLevel', 'ws.Columns(2).OutlineLevel = 3', 'ws.Columns("B:C")'),
    ('PageBreak', 'ws.Rows(2).PageBreak = xlPageBreakManual', 'ws.Rows("2:3")'),
    ('PageBreak', 'ws.Columns(2).PageBreak = xlPageBreakManual', 'ws.Columns("B:C")'),
    ('UseStandardHeight', 'ws.Rows(2).RowHeight = 40', 'ws.Rows("2:3")'),
    ('UseStandardWidth', 'ws.Columns(2).ColumnWidth = 30', 'ws.Columns("B:C")'),
    # With the second changed, which tells the first's answer from any's.
    ('Hidden', 'ws.Rows(3).Hidden = True', 'ws.Rows("2:3")'),
    ('Hidden', 'ws.Columns(3).Hidden = True', 'ws.Columns("B:C")'),
    ('OutlineLevel', 'ws.Rows(3).OutlineLevel = 3', 'ws.Rows("2:3")'),
    ('OutlineLevel', 'ws.Columns(3).OutlineLevel = 3', 'ws.Columns("B:C")'),
    ('PageBreak', 'ws.Rows(3).PageBreak = xlPageBreakManual', 'ws.Rows("2:3")'),
    ('PageBreak', 'ws.Columns(3).PageBreak = xlPageBreakManual', 'ws.Columns("B:C")'),
    ('UseStandardHeight', 'ws.Rows(3).RowHeight = 40', 'ws.Rows("2:3")'),
    ('UseStandardWidth', 'ws.Columns(3).ColumnWidth = 30', 'ws.Columns("B:C")'),
]


def differing(typed: bool = False) -> str:
    """A probe of each read over cells that differ, its answers joined by @@."""
    kind = 'Range' if typed else 'Object'
    lines = ['Public Function Probe() As String',
             f'    Dim wb As Object, ws As Object, c As {kind}, v As Variant, s As String',
             '    Set wb = Workbooks.Add(xlWBATWorksheet)']
    for member, setup, target in DIFFERING:
        lines += ['    Set ws = wb.Worksheets.Add', '    On Error Resume Next', f'    {setup}', '    Err.Clear',
                  f'    Set c = {target}', f'    v = "unread": v = c.{member}', '    If Err.Number <> 0 Then',
                  '        s = s & "@@e" & Err.Number', '        Err.Clear', '    ElseIf IsNull(v) Then',
                  '        s = s & "@@Null"', '    ElseIf IsArray(v) Then', '        s = s & "@@" & TypeName(v)',
                  '    Else', '        s = s & "@@" & TypeName(v) & " " & CStr(v)', '    End If', '    On Error GoTo 0']
    return '\n'.join([*lines, '    wb.Close False', '    Probe = Mid(s, 3)', 'End Function', ''])


def main() -> None:
    more.MEMBERS = BLOCKS
    more.NEEDS_96_DPI = False
    more.OUT = OUT
    more.main()
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0, exclusive=False)) as excel:
        excel.new_document()
        late = excel.run_vba(differing(), 'Probe', (), timeout=120.0)
        assert late.ok and not late.dialogs, late
        early = excel.run_vba(differing(typed=True), 'Probe', (), timeout=120.0)
        assert early.ok and not early.dialogs, early
    record = json.loads(OUT.read_text(encoding='utf-8'))
    record['differing_source'] = differing()
    record['differing'] = [[*one, answer, typed] for one, answer, typed in
                           zip(DIFFERING, str(late.value).split('@@'), str(early.value).split('@@'), strict=True)]
    OUT.write_text(json.dumps(record, indent=0) + '\n', encoding='utf-8')
    for one in record['differing']:
        print(one[0], one[2], '|', one[3], '|', one[4])


if __name__ == '__main__':
    main()
