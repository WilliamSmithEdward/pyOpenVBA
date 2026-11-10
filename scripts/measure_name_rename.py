"""Name round 55101's three rules, measured apart: a name renamed to its own spelling in another case, a name renamed to a text
holding "!", and ValidWorkbookParameter on a book saved and reopened. Rules written before this probe:

Case. A. The name takes the new case (what the model did). B. Nothing changes and no error. C. Error 1004, nothing changes.
Bang. A. The text is a qualified name: a sheet's name and "!" make the name that sheet's own, as Names.Add does, and a sheet
   that is not there is 1004 (what the model did). B. The whole text is the name's bare part, so a sheet's own name reads with
   its sheet's name twice ("Data!Data!Loc"), and a workbook's name reads as given. C. Any "!" is 1004.
Valid. A. True for every saved one-cell $A$1 reference (what the model did). B. False for a cell in a pivot table's source.
   C. False for the first row of a pivot table's source only. D. False once the book is reopened, until it is edited.

Each case runs on a new workbook, typed and late bound, reading Err straight after the statement; by live Excel and the model,
written to tests/fixtures/name_rename.json.

    python scripts/measure_name_rename.py
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
OUT = ROOT / 'tests' / 'fixtures' / 'name_rename.json'
WB = 'Set n = wb.Names.Add("Abc", "=1")'
LOCAL = 'Set n = ws2.Names.Add("Loc", "=2")'
#: Each case: its group, the statements that make n, and the one statement measured.
RENAMES = [
    ('case', WB, 'n.Name = "ABC"'), ('case', WB, 'n.Name = "abc"'), ('case', WB, 'n.Name = "Abc"'), ('case', WB, 'n.Name = "Abd"'),
    ('case', LOCAL, 'n.Name = "LOC"'), ('case', LOCAL, 'n.Name = "Lod"'),
    ('case', WB + ': wb.Names.Add "Xyz", "=4": Set n = wb.Names("Xyz")', 'n.Name = "ABC"'),
    ('bang', LOCAL, 'n.Name = "Data!Loc"'), ('bang', LOCAL, 'n.Name = "Data!Other"'), ('bang', LOCAL, 'n.Name = "Sheet1!Other"'),
    ('bang', WB, 'n.Name = "Data!Abc"'), ('bang', WB, 'n.Name = "Data!New"'), ('bang', WB, 'n.Name = "Nosheet!New"'),
    ('bang', WB, 'n.Name = "A!B!C"'), ('bang', WB, 'n.Name = "!New"'), ('bang', WB, 'n.Name = "New!"'),
    # Second probe: which sheet a qualifier may name, with the sheet in front varied, and whether the result is found.
    ('bang', WB, 'n.Name = "Sheet1!New"'), ('bang', WB + ': ws.Activate', 'n.Name = "Sheet1!New"'), ('bang', WB + ': ws.Activate', 'n.Name = "Data!New"'),
    ('bang', LOCAL + ': ws.Activate', 'n.Name = "Data!Other"'), ('bang', LOCAL + ': ws.Activate', 'n.Name = "Sheet1!Other"'),
    ('bang', 'Set n = ws.Names.Add("Mine", "=3")', 'n.Name = "Sheet1!Other"'), ('bang', 'Set n = ws.Names.Add("Mine", "=3")', 'n.Name = "Data!Other"'),
    ('bang', LOCAL, 'n.Name = "Data!Other": s = TypeName(ws2.Names("Data!Other")) & "/" & TypeName(wb.Names("Data!Data!Other")) & "/" & ws2.Evaluate("Data!Other")'),
    ('bang', WB, 'n.Name = "Data!New": s = TypeName(wb.Names("Data!New")) & "/" & wb.Names("Data!New").Parent.Worksheets.Count & "/" & ws.Evaluate("Data!New")'),
    ('bang', LOCAL, 'n.Name = "data!Other"'), ('bang', LOCAL, 'n.Name = "\'Data\'!Other"'), ('bang', WB, 'n.Name = "Data!1bad"'),
]
READS = ('    s = s & " || " & e & " | " & n.Name & " | " & n.NameLocal & " | " & TypeName(n.Parent) & " | " & n.Index\n'
         '    For i = 1 To wb.Names.Count: s = s & " | " & wb.Names(i).Name: Next\n'
         '    s = s & " | " & ws2.Names.Count & " E" & Err.Number\n')
DATA = '    ws.Range("A1:B3").Value = Array("k", "v"): ws.Range("A2:B2").Value = Array("a", 1): ws.Range("A3:B3").Value = Array("b", 2)\n'
PIVOT = '    wb.PivotCaches.Create(xlDatabase, ws.Range("A1:B3")).CreatePivotTable ws2.Range("A3"), "PT"\n'
REOPEN = ('    wb.SaveAs path & "\\v" & Format(Now, "hhnnss") & (Int(Timer * 1000) Mod 1000) & ".xlsx", xlOpenXMLWorkbook\n'
          '    Set wb = Workbooks.Open(wb.FullName): Set ws = wb.Worksheets(1): Set ws2 = wb.Worksheets(2)\n')
SAVE = '    wb.SaveAs path & "\\s" & Format(Now, "hhnnss") & (Int(Timer * 1000) Mod 1000) & ".xlsx", xlOpenXMLWorkbook\n'
NAMES = '    wb.Names.Add "Pa", "=Sheet1!$A$1": wb.Names.Add "Pb", "=Sheet1!$A$2": wb.Names.Add "Pc", "=Sheet1!$D$5"\n'
#: Each case: the book's making, then the reads of the three names' ValidWorkbookParameter.
VALIDS = [
    ('names saved', NAMES + SAVE), ('names reopened', NAMES + REOPEN), ('data saved', DATA + NAMES + SAVE),
    ('data reopened', DATA + NAMES + REOPEN), ('pivot saved', DATA + PIVOT + NAMES + SAVE), ('pivot reopened', DATA + PIVOT + NAMES + REOPEN),
    ('pivot reopened, column inserted', DATA + PIVOT + NAMES + REOPEN + '    ws.Columns(1).Insert\n'),
    ('pivot reopened, a cell set', DATA + PIVOT + NAMES + REOPEN + '    ws.Range("H9").Value = 1\n'),
    ('table reopened', DATA + '    ws.ListObjects.Add xlSrcRange, ws.Range("F1:G3"), , xlYes\n' + NAMES + REOPEN),
    # Second probe, after a reopen. Rules: D1 True once that name's reference moves; D2 True after any row or column insert
    # or delete in the book; D3 True once the book is saved again; D4 a name added after the reopen is True.
    ('reopened, rows inserted', NAMES + REOPEN + '    ws.Rows(1).Insert\n'), ('reopened, column Z inserted', NAMES + REOPEN + '    ws.Columns("Z").Insert\n'),
    ('reopened, other sheet column inserted', NAMES + REOPEN + '    ws2.Columns(1).Insert\n'), ('reopened, row 9 deleted', NAMES + REOPEN + '    ws.Rows(9).Delete\n'),
    ('reopened, saved again', NAMES + REOPEN + '    wb.Save\n'), ('reopened, name added', NAMES + REOPEN + '    wb.Names.Add "Pd", "=Sheet1!$C$3"\n'),
    ('reopened, RefersTo set', NAMES + REOPEN + '    wb.Names("Pa").RefersTo = "=Sheet1!$A$1"\n'),
    ('reopened, A1 set', NAMES + REOPEN + '    ws.Range("A1").Value = 5\n'),
    # Third probe: edits that move cells without a whole row or column, and others that change the book's structure.
    # Rules: E1 any cell shift counts as an insert does; E2 only whole rows and columns count; sheets added, moved or
    # deleted count under neither.
    ('reopened, cells shifted down', NAMES + REOPEN + '    ws.Range("C3").Insert xlShiftDown\n'),
    ('reopened, cells shifted left', NAMES + REOPEN + '    ws.Range("F3:F4").Delete xlShiftToLeft\n'),
    ('reopened, sheet added', NAMES + REOPEN + '    wb.Worksheets.Add\n'), ('reopened, sheet deleted', NAMES + REOPEN + '    ws2.Delete\n'),
    ('reopened, cells cut', NAMES + REOPEN + '    ws.Range("H1").Value = 1: ws.Range("H1").Cut ws.Range("J1")\n'),
    ('new, rows inserted', NAMES + '    ws.Rows(1).Insert\n'), ('new, rows inserted then saved', NAMES + '    ws.Rows(1).Insert\n' + SAVE),
    # Fourth probe: whose flag. Rules: F1 the book's, cleared by Workbooks.Open of it; F2 the application's, cleared by
    # any Workbooks.Open; F3 cleared by any workbook made or opened.
    ('saved, another book opened', NAMES + SAVE + '    Workbooks.Open(path & "\\t.xlsx").Close False\n'),
    ('saved, another book added', NAMES + SAVE + '    Workbooks.Add.Close False\n'),
    ('saved, closed and opened', NAMES + SAVE + '    i = 0: s = wb.FullName: wb.Close False: Set wb = Workbooks.Open(s): Set ws = wb.Worksheets(1): Set ws2 = wb.Worksheets(2)\n'),
    ('saved, closed, opened, rows inserted', NAMES + SAVE + '    s = wb.FullName: wb.Close False: Set wb = Workbooks.Open(s): Set ws = wb.Worksheets(1): Set ws2 = wb.Worksheets(2)\n'
     '    ws.Rows(1).Insert\n'),
]
VALID_READS = ('    s = ""\n    For i = 1 To wb.Names.Count\n        s = s & wb.Names(i).Name & "=" & wb.Names(i).RefersTo & ":" & wb.Names(i).ValidWorkbookParameter & " "\n'
               '    Next\n    s = s & "E" & Err.Number\n')


def source(case: dict[str, str], typed: bool) -> str:
    kind = 'Name' if typed else 'Object'
    if case['group'] == 'round':
        return (rounds.HELPER + 'Public Function One(ByVal path As String) As String\n' + rounds.DIMS.format(kind=kind)
                + '    Dim s As String, i As Long\n    Application.DisplayAlerts = False\n' + rounds.setup('reopened', 55101)
                + '    On Error Resume Next\n    Set o = wbL.Names(1)\n    s = o.Name & ":" & o.ValidWorkbookParameter & " " & wbL.Names(2).Name & ":" & wbL.Names(2).ValidWorkbookParameter\n'
                + '    s = s & " E" & Err.Number\n    On Error GoTo 0\n    wbT.Close False\n    One = s\nEnd Function\n')
    head = ('Public Function One(ByVal path As String) As String\n'
            f'    Dim wb As Object, ws As Object, ws2 As Object, n As {kind}, s As String, e As Long, i As Long\n    Application.DisplayAlerts = False\n'
            '    Set wb = Workbooks.Add(xlWBATWorksheet): Set ws = wb.Worksheets(1)\n'
            '    Set ws2 = wb.Worksheets.Add(After:=ws): ws2.Name = "Data"\n    On Error Resume Next\n')
    if case['group'] == 'valid':
        body = case['make'] + VALID_READS
    else:
        body = f'    {case["make"]}\n    Err.Clear\n    {case["call"]}\n    e = Err.Number: Err.Clear\n' + READS
    return head + body + '    On Error GoTo 0\n    wb.Close False\n    One = s\nEnd Function\n'


def model(case: dict[str, str], typed: bool, folder: str) -> str:
    from pyopenvba.apps.excel import ExcelApplication
    from pyopenvba.exceptions import VBAError

    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(source(case, typed), name='Rename')
    try:
        return str(app.run('Rename.One', folder))
    except VBAError as failed:
        return f'model {type(failed).__name__}: {str(failed)[:120]}'


def cases() -> list[dict[str, str]]:
    made = [{'group': group, 'make': make, 'call': call} for group, make, call in RENAMES]
    # NameLocal set to the same texts, whose refusals are 1004 where Name's are silent (tests/fixtures/settable_lists.json).
    made += [{'group': group, 'make': make, 'call': call.replace('n.Name =', 'n.NameLocal =')} for group, make, call in RENAMES if ': s =' not in call]
    made += [{'group': 'valid', 'make': make, 'call': label} for label, make in VALIDS]
    return made + [{'group': 'round', 'make': '', 'call': 'round 55101 reopened book'}]


def _lock() -> Path:
    common = subprocess.run(['git', 'rev-parse', '--git-common-dir'], capture_output=True, text=True, check=True, cwd=ROOT).stdout.strip()
    return (ROOT / common).resolve() / 'excel-coordination' / 'live.lock'


def main() -> None:
    from pyvbaharness import ExcelSession, HarnessConfig

    sys.stdout.reconfigure(encoding='utf-8')  # type: ignore[union-attr]
    made = cases()
    with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as held:
        rounds.books(Path(held))
        lock = _lock()
        owner = rounds.take_lock(lock, 'Verification 02')
        try:
            with ExcelSession(HarnessConfig(lock_wait_s=1200.0, exclusive=False)) as excel:
                excel.new_document()
                for case in made:
                    for tag, typed in (('late', False), ('typed', True)):
                        done = excel.run_vba(source(case, typed), 'One', (held,), timeout=60.0, line_numbers=False)
                        case[tag] = str(done.value) if done.ok else f'no answer: {str(done)[:300]}'
        finally:
            rounds.release_lock(lock, owner)
        for case in made:
            for tag, typed in (('late', False), ('typed', True)):
                case[f'model_{tag}'] = model(case, typed, held)
    OUT.write_text(json.dumps({'cases': made}, indent=1, ensure_ascii=False) + '\n', encoding='utf-8')
    for case in made:
        print(case['group'], case['call'], '| Excel', case['late'], '| typed same' if case['typed'] == case['late'] else f'| typed {case["typed"]}',
              '| model', case['model_late'])


if __name__ == '__main__':
    main()
