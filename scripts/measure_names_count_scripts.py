"""Names.Count with names in scripts the model cannot sort: CJK, kana, Hangul, Arabic and Thai beside Latin.

The rule written before Excel answered: Count is every name the collection holds, in any script, whatever their order would be;
a sheet's Names counts the sheet's own. Each name is added (an Add Excel refuses is recorded and left out), then the counts are
read late bound and typed on the new book and on the book saved and reopened, by Excel live and by the model, and written to
tests/fixtures/names_count_scripts.json.

    python scripts/measure_names_count_scripts.py
"""
from __future__ import annotations

import json
import os
import subprocess
import tempfile
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / 'tests' / 'fixtures' / 'names_count_scripts.json'
BOOK_NAMES = ['Ab', '名前', 'なまえ', '이름', 'اسم', 'ชื่อ', 'カード']
SHEET_NAMES = ['ローカル', 'Zed']


def _vba(text: str) -> str:
    """A VBA string expression for text, with ChrW for anything beyond ASCII."""
    return ' & '.join(f'ChrW({ord(ch)})' if ord(ch) > 127 else f'"{ch}"' for ch in text)


def source(typed: bool, reopen: bool) -> str:
    kind = 'Names' if typed else 'Object'
    lines = ['Public Function One(ByVal path As String) As String', f'    Dim wb As Object, ws As Object, n As {kind}, m As {kind}, s As String, saved As String',
             '    Application.DisplayAlerts = False', '    Set wb = Workbooks.Add(xlWBATWorksheet): Set ws = wb.Worksheets(1)', '    On Error Resume Next']
    for at, name in enumerate(BOOK_NAMES):
        lines += [f'    wb.Names.Add {_vba(name)}, "=Sheet1!$A${at + 1}"', '    s = s & " b" & Err.Number: Err.Clear']
    for at, name in enumerate(SHEET_NAMES):
        lines += [f'    ws.Names.Add {_vba(name)}, "=Sheet1!$B${at + 1}"', '    s = s & " s" & Err.Number: Err.Clear']
    if reopen:
        lines += ['    saved = path & "\\n" & Format(Now, "hhnnss") & (Int(Timer * 1000) Mod 1000) & ".xlsx"', '    wb.SaveAs saved, xlOpenXMLWorkbook',
                  '    wb.Close False', '    Set wb = Workbooks.Open(saved): Set ws = wb.Worksheets(1)']
    lines += ['    Set n = wb.Names: Set m = ws.Names', '    s = s & " | " & n.Count & " " & TypeName(n.Count) & " E" & Err.Number: Err.Clear',
              '    s = s & " | " & m.Count & " " & TypeName(m.Count) & " E" & Err.Number: Err.Clear', '    On Error GoTo 0', '    wb.Close False',
              '    One = Mid(s, 2)', 'End Function']
    return '\n'.join(lines) + '\n'


def model(typed: bool, reopen: bool, folder: str) -> str:
    from pyopenvba.apps.excel import ExcelApplication
    from pyopenvba.exceptions import VBAError

    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(source(typed, reopen), name='Count')
    try:
        return str(app.run('Count.One', folder))
    except VBAError as failed:
        return f'model {type(failed).__name__}: {str(failed)[:120]}'


def _lock() -> Path:
    common = subprocess.run(['git', 'rev-parse', '--git-common-dir'], capture_output=True, text=True, check=True, cwd=ROOT).stdout.strip()
    return (ROOT / common).resolve() / 'excel-coordination' / 'live.lock'


def main() -> None:
    from pyvbaharness import ExcelSession, HarnessConfig

    cases = [{'typed': typed, 'reopen': reopen} for typed in (False, True) for reopen in (False, True)]
    with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as held:
        for case in cases:
            case['model'] = model(case['typed'], case['reopen'], held)
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
                    done = excel.run_vba(source(case['typed'], case['reopen']), 'One', (held,), timeout=60.0, line_numbers=False)
                    case['excel'] = str(done.value) if done.ok else f'no answer: {str(done)[:300]}'
        finally:
            os.rmdir(lock)
    OUT.write_text(json.dumps({'book_names': BOOK_NAMES, 'sheet_names': SHEET_NAMES, 'cases': cases}, indent=1, ensure_ascii=False) + '\n', encoding='utf-8')
    for case in cases:
        print(case['typed'], case['reopen'], '| Excel', case['excel'], '| model', case['model'])


if __name__ == '__main__':
    main()
