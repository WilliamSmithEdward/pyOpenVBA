"""What a table's default member answers when its Name and DisplayName differ. ListObject round 53101 had Excel answer x = o
with the given name Tb42 where the model gave the DisplayName Tb42_. Rules written before this probe:

A. The default member is DisplayName (what the model did).
B. It is Name, the name given.

Each table is named, then read as x = o and CStr(o), beside Name and DisplayName, late bound and typed, on the new book and on
the book saved and reopened, by live Excel and the model, and written to tests/fixtures/list_object_default.json.

    python scripts/measure_list_object_default.py
"""
from __future__ import annotations

import json
import os
import subprocess
import tempfile
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / 'tests' / 'fixtures' / 'list_object_default.json'
NAMES = ['Tb7', 'a b', '1abc', 'Plain', 'XFD1']


def source(name: str, typed: bool, reopen: bool) -> str:
    kind = 'ListObject' if typed else 'Object'
    lines = ['Public Function One(ByVal path As String) As String', f'    Dim wb As Object, ws As Object, o As {kind}, s As String, x As String, saved As String',
             '    Application.DisplayAlerts = False', '    Set wb = Workbooks.Add(xlWBATWorksheet): Set ws = wb.Worksheets(1)',
             '    ws.Range("B2:C2").Value = Array("h", "g"): ws.Range("B3:C4").Value = 1',
             f'    ws.ListObjects.Add(xlSrcRange, ws.Range("B2:C4"), , xlYes).Name = "{name}"']
    if reopen:
        lines += ['    saved = path & "\\d" & Format(Now, "hhnnss") & (Int(Timer * 1000) Mod 1000) & ".xlsx"', '    wb.SaveAs saved, xlOpenXMLWorkbook',
                  '    wb.Close False', '    Set wb = Workbooks.Open(saved): Set ws = wb.Worksheets(1)']
    lines += ['    Set o = ws.ListObjects(1)', '    On Error Resume Next', '    x = o: s = x & " E" & Err.Number: Err.Clear: x = ""',
              '    s = s & " | " & CStr(o) & " E" & Err.Number: Err.Clear', '    s = s & " | " & o.Name & " | " & o.DisplayName',
              '    On Error GoTo 0', '    wb.Close False', '    One = s', 'End Function']
    return '\n'.join(lines) + '\n'


def model(name: str, typed: bool, reopen: bool, folder: str) -> str:
    from pyopenvba.apps.excel import ExcelApplication
    from pyopenvba.exceptions import VBAError

    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(source(name, typed, reopen), name='Default')
    try:
        return str(app.run('Default.One', folder))
    except VBAError as failed:
        return f'model {type(failed).__name__}: {str(failed)[:120]}'


def _lock() -> Path:
    common = subprocess.run(['git', 'rev-parse', '--git-common-dir'], capture_output=True, text=True, check=True, cwd=ROOT).stdout.strip()
    return (ROOT / common).resolve() / 'excel-coordination' / 'live.lock'


def main() -> None:
    from pyvbaharness import ExcelSession, HarnessConfig

    cases = [{'name': name, 'typed': typed, 'reopen': reopen} for name in NAMES for typed in (False, True) for reopen in (False, True)]
    with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as held:
        for case in cases:
            case['model'] = model(case['name'], case['typed'], case['reopen'], held)
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
                    done = excel.run_vba(source(case['name'], case['typed'], case['reopen']), 'One', (held,), timeout=60.0, line_numbers=False)
                    case['excel'] = str(done.value) if done.ok else f'no answer: {str(done)[:300]}'
        finally:
            os.rmdir(lock)
    for case in cases:
        parts = case['excel'].split(' | ')
        if len(parts) == 4:
            given, shown = parts[2], parts[3]
            case['rule_a'] = f'{shown} E0 | {shown} E0 | {given} | {shown}'
            case['rule_b'] = f'{given} E0 | {given} E0 | {given} | {shown}'
    OUT.write_text(json.dumps({'cases': cases}, indent=1) + '\n', encoding='utf-8')
    for rule in ('rule_a', 'rule_b', 'model'):
        print(rule, 'agrees on', sum(case['excel'] == case.get(rule) for case in cases), 'of', len(cases))
    for case in cases:
        print(case['name'], case['typed'], case['reopen'], '| Excel', case['excel'], '| model', case['model'])


if __name__ == '__main__':
    main()
