"""ListObject.TableStyle set by name on a book with a table style of its own: tests/fixtures/dxf_table/made.xlsx, which Excel saved
holding "Mine" (pivot="0" table="0"). ListObject round 53103 had Excel keep its style where the model took "NoSuchStyle". Rules
written before this probe:

A. Any name is taken once the book has a style of its own (what the model did).
B. A built-in name, or the book's own style's name in any case; anything else is error 450, as on a book without one.
C. As B, but a style the book marks table="0" is refused as an unknown one.

Each name is set on the book's table straight after it is opened, the error read before anything else, then the style read, by
live Excel and the model, written to tests/fixtures/table_style_own.json.

    python scripts/measure_table_style_own.py
"""
from __future__ import annotations

import json
import os
import shutil
import subprocess
import tempfile
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
BOOK = ROOT / 'tests' / 'fixtures' / 'dxf_table' / 'made.xlsx'
OUT = ROOT / 'tests' / 'fixtures' / 'table_style_own.json'
NAMES = ['NoSuchStyle', 'Mine', 'MINE', 'mine ', 'TableStyleLight2', 'tablestylelight2', '', 'PivotStyleLight16']


def source(name: str, typed: bool) -> str:
    kind = 'ListObject' if typed else 'Object'
    return ('Public Function One(ByVal path As String) As String\n'
            f'    Dim wb As Object, o As {kind}, s As String, e As Long\n    Application.DisplayAlerts = False\n'
            '    Set wb = Workbooks.Open(path & "\\made.xlsx")\n    Set o = wb.Worksheets(1).ListObjects(1)\n'
            '    On Error Resume Next\n'
            f'    o.TableStyle = "{name}"\n    e = Err.Number: Err.Clear\n'
            '    s = e & " | " & TypeName(o.TableStyle) & " " & o.TableStyle.Name & " E" & Err.Number\n    Err.Clear\n'
            '    On Error GoTo 0\n    wb.Close False\n    One = s\nEnd Function\n')


def model(name: str, typed: bool, folder: str) -> str:
    from pyopenvba.apps.excel import ExcelApplication
    from pyopenvba.exceptions import VBAError

    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(source(name, typed), name='Style')
    try:
        return str(app.run('Style.One', folder))
    except VBAError as failed:
        return f'model {type(failed).__name__}: {str(failed)[:120]}'


def _lock() -> Path:
    common = subprocess.run(['git', 'rev-parse', '--git-common-dir'], capture_output=True, text=True, check=True, cwd=ROOT).stdout.strip()
    return (ROOT / common).resolve() / 'excel-coordination' / 'live.lock'


def main() -> None:
    from pyvbaharness import ExcelSession, HarnessConfig

    cases = [{'name': name, 'typed': typed} for name in NAMES for typed in (False, True)]
    with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as held:
        shutil.copyfile(BOOK, Path(held) / 'made.xlsx')
        for case in cases:
            case['model'] = model(case['name'], case['typed'], held)
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
                    done = excel.run_vba(source(case['name'], case['typed']), 'One', (held,), timeout=60.0, line_numbers=False)
                    case['excel'] = str(done.value) if done.ok else f'no answer: {str(done)[:300]}'
        finally:
            os.rmdir(lock)
    OUT.write_text(json.dumps({'cases': cases}, indent=1) + '\n', encoding='utf-8')
    for case in cases:
        print(repr(case['name']), case['typed'], '| Excel', case['excel'], '| model', case['model'])


if __name__ == '__main__':
    main()
