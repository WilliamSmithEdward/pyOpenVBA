"""A table's AutoFilter.Range beside its totals row. ListObject round 53105 had Excel answer H1:I6 on a reopened book where the
model gave H1:I5. Rules written before this probe:

A. The header and data rows, never the totals row (what the model did).
B. Taking in the totals row once the table has come back from a file.
C. Taking in the totals row when totals were turned on before the filter was made, new or reopened.

Each case builds a table, sets totals and the filter in an order, reads AutoFilter.Range beside Range, new and saved and reopened by
live Excel and the model; Excel's saved autoFilter ref is read from the file. Written to tests/fixtures/table_filter_range.json.

    python scripts/measure_table_filter_range.py
"""
from __future__ import annotations

import json
import os
import re
import subprocess
import tempfile
import time
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / 'tests' / 'fixtures' / 'table_filter_range.json'
#: (what is done to the table after it is made, in order)
ORDERS = {
    'totals': ['o.ShowTotals = True'],
    'none': [],
    'filter off, totals, filter on': ['o.ShowAutoFilter = False', 'o.ShowTotals = True', 'o.ShowAutoFilter = True'],
    'totals, filter off, filter on': ['o.ShowTotals = True', 'o.ShowAutoFilter = False', 'o.ShowAutoFilter = True'],
    'totals, criteria': ['o.ShowTotals = True', 'o.Range.AutoFilter 1, "<5"'],
    'totals then off': ['o.ShowTotals = True', 'o.ShowTotals = False'],
    'criteria, totals': ['o.Range.AutoFilter 1, "<5"', 'o.ShowTotals = True'],
    'criteria, totals, criteria': ['o.Range.AutoFilter 1, "<5"', 'o.ShowTotals = True', 'o.Range.AutoFilter 2, "<9"'],
}


def source(steps: list[str], reopen: bool, typed: bool) -> str:
    kind = 'ListObject' if typed else 'Object'
    lines = ['Public Function One(ByVal path As String) As String', f'    Dim wb As Object, ws As Object, o As {kind}, s As String, saved As String',
             '    Application.DisplayAlerts = False', '    Set wb = Workbooks.Add(xlWBATWorksheet): Set ws = wb.Worksheets(1)',
             '    ws.Range("H1:I1").Value = Array("p", "q"): ws.Range("H2:I4").Value = 3', '    Set o = ws.ListObjects.Add(xlSrcRange, ws.Range("H1:I4"), , xlYes)',
             *[f'    {step}' for step in steps], '    saved = path & "\\f" & Format(Now, "hhnnss") & (Int(Timer * 1000) Mod 1000) & ".xlsx"']
    if reopen:
        lines += ['    wb.SaveAs saved, xlOpenXMLWorkbook', '    wb.Close False', '    Set wb = Workbooks.Open(saved): Set ws = wb.Worksheets(1)',
                  '    Set o = ws.ListObjects(1)']
    lines += ['    On Error Resume Next', '    s = o.Range.Address(False, False) & " | " & o.AutoFilter.Range.Address(False, False) & " E" & Err.Number',
              '    Err.Clear: On Error GoTo 0', '    If Not wb.Path = "" Then s = s & " | " & Mid(wb.FullName, Len(path) + 2)', '    wb.Close False', '    One = s', 'End Function']
    return '\n'.join(lines) + '\n'


def model(steps: list[str], reopen: bool, typed: bool, folder: str) -> str:
    from pyopenvba.apps.excel import ExcelApplication
    from pyopenvba.exceptions import VBAError

    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(source(steps, reopen, typed), name='Filter')
    try:
        return str(app.run('Filter.One', folder)).split(' | f')[0]
    except VBAError as failed:
        return f'model {type(failed).__name__}: {str(failed)[:120]}'


def _lock() -> Path:
    common = subprocess.run(['git', 'rev-parse', '--git-common-dir'], capture_output=True, text=True, check=True, cwd=ROOT).stdout.strip()
    return (ROOT / common).resolve() / 'excel-coordination' / 'live.lock'


def main() -> None:
    from pyvbaharness import ExcelSession, HarnessConfig

    cases = [{'order': order, 'reopen': reopen, 'typed': typed} for order in ORDERS for reopen in (False, True) for typed in (False, True)]
    with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as held:
        for case in cases:
            case['model'] = model(ORDERS[case['order']], case['reopen'], case['typed'], held)
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
                    done = excel.run_vba(source(ORDERS[case['order']], case['reopen'], case['typed']), 'One', (held,), timeout=60.0, line_numbers=False)
                    answer = str(done.value) if done.ok else f'no answer: {str(done)[:300]}'
                    parts = answer.split(' | ')
                    if len(parts) == 3:
                        with zipfile.ZipFile(Path(held) / parts[2]) as package:
                            table = next(name for name in package.namelist() if name.startswith('xl/tables/table'))
                            found = re.search(rb'<autoFilter ref="([^"]+)"', package.read(table))
                            case['file_ref'] = found.group(1).decode() if found else None
                        answer = ' | '.join(parts[:2])
                    case['excel'] = answer
        finally:
            os.rmdir(lock)
    OUT.write_text(json.dumps({'cases': cases}, indent=1) + '\n', encoding='utf-8')
    for case in cases:
        print(case['order'], case['reopen'], case['typed'], '| Excel', case['excel'], '| file', case.get('file_ref'), '| model', case['model'])


if __name__ == '__main__':
    main()
