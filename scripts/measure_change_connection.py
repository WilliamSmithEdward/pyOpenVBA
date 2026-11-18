"""PivotTable.ChangeConnection on a table over a list, by what it is given. PivotTable round 57101 had Excel raise 1004 for
Nothing through CallByName, where the model raised 13 (tests/fixtures/pivot_table_more.json measured 13 for a worksheet).
Rules written before this probe:

A. A connection is 5, anything else 13 (what the model did).
B. Nothing is 1004, whichever way it is passed; another object 13; a connection 5.
C. Every value that is not an object (text, a number, Empty, Null) is 1004 too, only an object of another kind 13.

Each argument is passed late bound, as a statement and through CallByName, with Err read straight after; by live Excel and
the model, written to tests/fixtures/change_connection.json.

    python scripts/measure_change_connection.py
"""
from __future__ import annotations

import json
import subprocess
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import measure_table_pivot_rounds as rounds  # noqa: E402

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / 'tests' / 'fixtures' / 'change_connection.json'
ARGUMENTS = ['Nothing', 'ws', 'ws.Range("A1")', 'wb', '"text"', '5', 'Empty', 'Null', 'True']
WAYS = ['o.ChangeConnection {arg}', 'CallByName o, "ChangeConnection", VbMethod, {arg}']


def source(argument: str, way: str) -> str:
    call = way.format(arg='v')
    setter = 'Set v = ' if argument in ('Nothing', 'ws', 'ws.Range("A1")', 'wb') else 'v = '
    return ('Public Function One(ByVal path As String) As String\n'
            '    Dim wb As Object, ws As Object, o As Object, v As Variant, e As Long\n    Application.DisplayAlerts = False\n'
            '    Set wb = Workbooks.Add(xlWBATWorksheet): Set ws = wb.Worksheets(1)\n'
            '    ws.Range("A1:B3").Value = Array("k", "v"): ws.Range("A2:B2").Value = Array("a", 1): ws.Range("A3:B3").Value = Array("b", 2)\n'
            '    Set o = wb.PivotCaches.Create(xlDatabase, ws.Range("A1:B3")).CreatePivotTable(ws.Range("E3"), "PT")\n'
            f'    {setter}{argument}\n'
            f'    On Error Resume Next\n    {call}\n'
            '    e = Err.Number\n    On Error GoTo 0\n    wb.Close False\n    One = CStr(e)\nEnd Function\n')


def model(argument: str, way: str) -> str:
    from pyopenvba.apps.excel import ExcelApplication
    from pyopenvba.exceptions import VBAError

    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(source(argument, way), name='Connection')
    try:
        return str(app.run('Connection.One', ''))
    except VBAError as failed:
        return f'model {type(failed).__name__}: {str(failed)[:120]}'


def _lock() -> Path:
    common = subprocess.run(['git', 'rev-parse', '--git-common-dir'], capture_output=True, text=True, check=True, cwd=ROOT).stdout.strip()
    return (ROOT / common).resolve() / 'excel-coordination' / 'live.lock'


def cases() -> list[dict[str, str]]:
    return [{'argument': argument, 'way': way} for argument in ARGUMENTS for way in WAYS]


def main() -> None:
    from pyvbaharness import ExcelSession, HarnessConfig

    made = cases()
    with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as held:
        lock = _lock()
        owner = rounds.take_lock(lock, 'Verification 02')
        try:
            with ExcelSession(HarnessConfig(lock_wait_s=1200.0, exclusive=False)) as excel:
                excel.new_document()
                for case in made:
                    done = excel.run_vba(source(case['argument'], case['way']), 'One', (held,), timeout=60.0, line_numbers=False)
                    case['excel'] = str(done.value) if done.ok else f'no answer: {str(done)[:300]}'
        finally:
            rounds.release_lock(lock, owner)
    for case in made:
        case['model'] = model(case['argument'], case['way'])
    OUT.write_text(json.dumps({'cases': made}, indent=1) + '\n', encoding='utf-8')
    for case in made:
        print('same' if case['excel'] == case['model'] else 'DIFF', case['way'].format(arg=case['argument']), '| Excel', case['excel'], '| model', case['model'])


if __name__ == '__main__':
    main()
