"""Whether a totals table's AutoFilter.Range follows the active sheet. tests/fixtures/table_filter_block.json had the range whole while
the table's sheet was active and without the totals row once another was. The rule written before this probe:

R. A table's AutoFilter.Range is the whole table while its sheet is its workbook's active sheet, and the table less its totals row
   otherwise; a table without totals is whole either way. New and reopened books alike.

A book with a table on "T" and a sheet "Other" is read with T or Other active, new or saved and reopened, with totals or without,
and with another workbook in front; written to tests/fixtures/table_filter_active.json, live Excel beside the model.

    python scripts/measure_table_filter_active.py
"""
from __future__ import annotations

import json
import os
import subprocess
import tempfile
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / 'tests' / 'fixtures' / 'table_filter_active.json'
WHERE = {'T': 'wb.Worksheets("T").Activate', 'Other': 'wb.Worksheets("Other").Activate',
         'T, another book in front': 'wb.Worksheets("T").Activate: Set other = Workbooks.Add',
         'Other, another book in front': 'wb.Worksheets("Other").Activate: Set other = Workbooks.Add'}


def source(where: str, reopen: bool, totals: bool, criteria: bool, typed: bool) -> str:
    kind = 'ListObject' if typed else 'Object'
    lines = ['Public Function One(ByVal path As String) As String',
             f'    Dim wb As Object, ws As Object, o As {kind}, s As String, saved As String, other As Object',
             '    Application.DisplayAlerts = False', '    Set wb = Workbooks.Add(xlWBATWorksheet): Set ws = wb.Worksheets(1): ws.Name = "T"',
             '    wb.Worksheets.Add(After:=ws).Name = "Other"', '    ws.Activate',
             '    ws.Range("A1:B1").Value = Array("p", "q"): ws.Range("A2:B2").Value = Array(1, 7): ws.Range("A3:B3").Value = Array(9, 2): ws.Range("A4:B4").Value = Array(3, 3)',
             '    Set o = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:B4"), , xlYes)']
    if criteria:
        lines.append('    o.Range.AutoFilter 1, "<5"')
    if totals:
        lines.append('    o.ShowTotals = True')
    if reopen:
        lines += ['    saved = path & "\\a" & Format(Now, "hhnnss") & (Int(Timer * 1000) Mod 1000) & ".xlsx"', '    wb.SaveAs saved, xlOpenXMLWorkbook', '    wb.Close False',
                  '    Set wb = Workbooks.Open(saved): Set ws = wb.Worksheets("T"): Set o = ws.ListObjects(1)']
    lines += [f'    {WHERE[where]}', '    On Error Resume Next', '    s = o.AutoFilter.Range.Address(False, False) & " E" & Err.Number',
              '    Err.Clear: On Error GoTo 0', '    If Not other Is Nothing Then other.Close False', '    wb.Close False', '    One = s', 'End Function']
    return '\n'.join(lines) + '\n'


def model(where: str, reopen: bool, totals: bool, criteria: bool, typed: bool, folder: str) -> str:
    from pyopenvba.apps.excel import ExcelApplication
    from pyopenvba.exceptions import VBAError

    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(source(where, reopen, totals, criteria, typed), name='Active')
    try:
        return str(app.run('Active.One', folder))
    except VBAError as failed:
        return f'model {type(failed).__name__}: {str(failed)[:120]}'


def rule(where: str, totals: bool) -> str:
    return ('A1:B4' if totals and not where.startswith('T') else 'A1:B5' if totals else 'A1:B4') + ' E0'


def _lock() -> Path:
    common = subprocess.run(['git', 'rev-parse', '--git-common-dir'], capture_output=True, text=True, check=True, cwd=ROOT).stdout.strip()
    return (ROOT / common).resolve() / 'excel-coordination' / 'live.lock'


def main() -> None:
    from pyvbaharness import ExcelSession, HarnessConfig

    cases = [{'where': where, 'reopen': reopen, 'totals': totals, 'criteria': criteria, 'typed': typed, 'rule': rule(where, totals)}
             for where in WHERE for reopen in (False, True) for totals in (True, False) for criteria in (True, False) for typed in (False, True)]
    with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as held:
        for case in cases:
            case['model'] = model(case['where'], case['reopen'], case['totals'], case['criteria'], case['typed'], held)
        lock = _lock()
        while True:
            try:
                os.mkdir(lock)
                break
            except FileExistsError:
                time.sleep(10)
        try:
            with ExcelSession(HarnessConfig(lock_wait_s=1200.0, exclusive=False)) as excel:
                excel.new_document()
                for case in cases:
                    done = excel.run_vba(source(case['where'], case['reopen'], case['totals'], case['criteria'], case['typed']), 'One', (held,),
                                         timeout=60.0, line_numbers=False)
                    case['excel'] = str(done.value) if done.ok else f'no answer: {str(done)[:300]}'
        finally:
            os.rmdir(lock)
    OUT.write_text(json.dumps({'cases': cases}, indent=1) + '\n', encoding='utf-8')
    print('rule R agrees on', sum(case['excel'] == case['rule'] for case in cases), 'of', len(cases))
    print('model agrees on', sum(case['excel'] == case['model'] for case in cases), 'of', len(cases))
    for case in cases:
        if case['excel'] != case['rule'] or case['excel'] != case['model']:
            print(case['where'], case['reopen'], case['totals'], case['criteria'], case['typed'], '| Excel', case['excel'], '| R', case['rule'], '| model', case['model'])


if __name__ == '__main__':
    main()
