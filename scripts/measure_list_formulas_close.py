"""PivotTable.ListFormulas and the workbook's Close after it. PivotTable round 57102 had Excel fail the book's Close with
-2147417848 after ListFormulas on the pivot table Excel saved in tests/fixtures/dxf_table/made.xlsx, once a column went in on its
source sheet; the model closed it. Rules written before this probe:

A. Close never fails after ListFormulas (what the model did).
B. It fails after ListFormulas on a table read from a file, whatever was edited.
C. It fails only after a column or row went in on the table's source sheet first, on any book.
D. It fails once ListFormulas has run on any table, the sheet it adds being the cause; ListFormulas itself answers no error.

Each case builds or opens the book, maybe inserts, runs ListFormulas reading Err, then closes the book reading Err, inside the
macro; by live Excel and the model, written to tests/fixtures/list_formulas_close.json.

    python scripts/measure_list_formulas_close.py
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
OUT = ROOT / 'tests' / 'fixtures' / 'list_formulas_close.json'
BOOKS = {
    'file': '    Set wbT = Workbooks.Open(path & "\\t.xlsx"): Set ws = wbT.Worksheets("one"): Set ws2 = wbT.Worksheets("pivot")\n',
    'new': ('    Set wbT = Workbooks.Add(xlWBATWorksheet): Set ws = wbT.Worksheets(1): ws.Name = "Main"\n'
            '    Set ws2 = wbT.Worksheets.Add(After:=ws): ws2.Name = "Second"\n'
            '    ws.Range("A1:B3").Value = Array("k", "v"): ws.Range("A2:B2").Value = Array("a", 1): ws.Range("A3:B3").Value = Array("b", 2)\n'
            '    wbT.PivotCaches.Create(xlDatabase, ws.Range("A1:B3")).CreatePivotTable ws2.Range("A3"), "PT"\n'),
}
#: Second look: the round's file book opens tests/fixtures/object_roundtrip.xlsx after it, which is then the book in front.
#: E. Close fails when another book was in front as ListFormulas ran; ListFormulas puts its sheet in that book.
BOOKS['file, another in front'] = BOOKS['file'] + '    Set wbL = Workbooks.Open(path & "\\l.xlsx")\n'
EDITS = ['', 'ws.Columns(1).Insert', 'ws.Rows(1).Insert', 'ws2.Columns(1).Insert']


def cases() -> list[dict[str, str]]:
    return [{'book': book, 'edit': edit, 'list': listed} for book in BOOKS for edit in EDITS for listed in ('True', 'False')]


def source(case: dict[str, str]) -> str:
    listing = ('    ws2.PivotTables(1).ListFormulas\n' if case['list'] == 'True' else '')
    return ('Public Function One(ByVal path As String) As String\n'
            '    Dim wbT As Object, wbL As Object, ws As Object, ws2 As Object, s As String\n    Application.DisplayAlerts = False\n'
            + BOOKS[case['book']] + '    On Error Resume Next\n' + (f'    {case["edit"]}\n' if case['edit'] else '')
            + '    s = "edit " & Err.Number: Err.Clear\n' + listing + '    s = s & " list " & Err.Number & " " & wbT.Worksheets.Count & " " & (ActiveWorkbook Is wbT): Err.Clear\n'
            '    If Not wbL Is Nothing Then s = s & " other " & wbL.Worksheets.Count: wbL.Close False: s = s & " " & Err.Number: Err.Clear\n'
            '    wbT.Close False\n    s = s & " close " & Err.Number\n    On Error GoTo 0\n    One = s\nEnd Function\n')


def model(case: dict[str, str], folder: str) -> str:
    from pyopenvba.apps.excel import ExcelApplication
    from pyopenvba.exceptions import VBAError

    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(source(case), name='Lists')
    try:
        return str(app.run('Lists.One', folder))
    except VBAError as failed:
        return f'model {type(failed).__name__}: {str(failed)[:120]}'


def _lock() -> Path:
    common = subprocess.run(['git', 'rev-parse', '--git-common-dir'], capture_output=True, text=True, check=True, cwd=ROOT).stdout.strip()
    return (ROOT / common).resolve() / 'excel-coordination' / 'live.lock'


def main() -> None:
    from pyvbaharness import ExcelSession, HarnessConfig

    made = cases()
    with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as held:
        rounds.books(Path(held))
        lock = _lock()
        owner = rounds.take_lock(lock, 'Verification 02')
        try:
            for case in made:
                # A new Excel each case: a failed Close can leave its book open.
                with ExcelSession(HarnessConfig(lock_wait_s=1200.0, exclusive=False)) as excel:
                    excel.new_document()
                    done = excel.run_vba(source(case), 'One', (held,), timeout=60.0, line_numbers=False)
                    case['excel'] = str(done.value) if done.ok else f'no answer: {str(done.error)[:200]}'
        finally:
            rounds.release_lock(lock, owner)
        for case in made:
            case['model'] = model(case, held)
    OUT.write_text(json.dumps({'cases': made}, indent=1) + '\n', encoding='utf-8')
    for case in made:
        print('same' if case['excel'] == case['model'] else 'DIFF', case['book'], repr(case['edit']), 'list', case['list'], '| Excel', case['excel'], '| model', case['model'])


if __name__ == '__main__':
    main()
