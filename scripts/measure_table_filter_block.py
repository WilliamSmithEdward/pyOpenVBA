"""When a table's filter stops answering the block its file gave. Excel writes a totals table's filter block without the totals row;
read back on a sheet not yet activated, AutoFilter.Range is that block (tests/fixtures/table_autofilter/, totals), and on the only
sheet, active as the book opens, it is the whole table (tests/fixtures/table_filter_range.json). The rule written before this probe:

H. The block lasts until the sheet's filter is settled as Excel settles FilterMode, when the sheet is activated, its selection moves
   or a filter changes; a sheet active as the book opens is settled then.

A book with the table on "T" (criteria, then totals) and a sheet "Other" is saved with either sheet active and opened again; T's
FilterMode and its table's AutoFilter.Range are read straight away, then after a step. Written to
tests/fixtures/table_filter_block.json, live Excel beside the model.

    python scripts/measure_table_filter_block.py
"""
from __future__ import annotations

import json
import os
import subprocess
import tempfile
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / 'tests' / 'fixtures' / 'table_filter_block.json'
STEPS = {'none': '', 'activate T': 'ws.Activate', 'activate Other': 'wb.Worksheets("Other").Activate',
         'select on T': 'ws.Activate: ws.Range("C2").Select', 'read only Range again': 'x = o.Range.Address'}


def source(active: str, step: str, typed: bool) -> str:
    kind = 'ListObject' if typed else 'Object'
    return '\n'.join([
        'Public Function One(ByVal path As String) As String',
        f'    Dim wb As Object, ws As Object, o As {kind}, s As String, saved As String, x As String',
        '    Application.DisplayAlerts = False', '    Set wb = Workbooks.Add(xlWBATWorksheet): Set ws = wb.Worksheets(1): ws.Name = "T"',
        '    wb.Worksheets.Add(After:=ws).Name = "Other"', '    ws.Activate',
        '    ws.Range("A1:B1").Value = Array("p", "q"): ws.Range("A2:B2").Value = Array(1, 7): ws.Range("A3:B3").Value = Array(9, 2): ws.Range("A4:B4").Value = Array(3, 3)',
        '    Set o = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:B4"), , xlYes)', '    o.Range.AutoFilter 1, "<5"', '    o.ShowTotals = True',
        f'    wb.Worksheets("{active}").Activate',
        '    saved = path & "\\b" & Format(Now, "hhnnss") & (Int(Timer * 1000) Mod 1000) & ".xlsx"', '    wb.SaveAs saved, xlOpenXMLWorkbook', '    wb.Close False',
        '    Set wb = Workbooks.Open(saved): Set ws = wb.Worksheets("T"): Set o = ws.ListObjects(1)',
        '    On Error Resume Next',
        '    s = ws.FilterMode & " " & o.AutoFilter.Range.Address(False, False)',
        f'    {STEPS[step]}' if STEPS[step] else '',
        '    s = s & " | " & ws.FilterMode & " " & o.AutoFilter.Range.Address(False, False) & " E" & Err.Number',
        '    Err.Clear: On Error GoTo 0', '    wb.Close False', '    One = s', 'End Function']) + '\n'


def model(active: str, step: str, typed: bool, folder: str) -> str:
    from pyopenvba.apps.excel import ExcelApplication
    from pyopenvba.exceptions import VBAError

    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(source(active, step, typed), name='Block')
    try:
        return str(app.run('Block.One', folder))
    except VBAError as failed:
        return f'model {type(failed).__name__}: {str(failed)[:120]}'


def rule(active: str, step: str) -> str:
    """H's prediction: the block A1:B4 until T is settled, the whole table A1:B5 after; FilterMode True once settled."""
    settled_at_open = active == 'T'
    first = 'True A1:B5' if settled_at_open else 'False A1:B4'
    settles = settled_at_open or step in ('activate T', 'select on T')
    return f'{first} | {"True A1:B5" if settles else "False A1:B4"} E0'


def _lock() -> Path:
    common = subprocess.run(['git', 'rev-parse', '--git-common-dir'], capture_output=True, text=True, check=True, cwd=ROOT).stdout.strip()
    return (ROOT / common).resolve() / 'excel-coordination' / 'live.lock'


def main() -> None:
    from pyvbaharness import ExcelSession, HarnessConfig

    cases = [{'active': active, 'step': step, 'typed': typed, 'rule': rule(active, step)}
             for active in ('T', 'Other') for step in STEPS for typed in (False, True)]
    with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as held:
        for case in cases:
            case['model'] = model(case['active'], case['step'], case['typed'], held)
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
                    done = excel.run_vba(source(case['active'], case['step'], case['typed']), 'One', (held,), timeout=60.0, line_numbers=False)
                    case['excel'] = str(done.value) if done.ok else f'no answer: {str(done)[:300]}'
        finally:
            os.rmdir(lock)
    OUT.write_text(json.dumps({'cases': cases}, indent=1) + '\n', encoding='utf-8')
    print('rule H agrees on', sum(case['excel'] == case['rule'] for case in cases), 'of', len(cases))
    for case in cases:
        print(case['active'], case['step'], case['typed'], '| Excel', case['excel'], '| H', case['rule'], '| model', case['model'])


if __name__ == '__main__':
    main()
