"""Held-out rounds over the VERIFIED rows of tables, pivot tables and the held families: Application, Creator, Parent, Count, Item
and the other members those rows name. A round is seeded. The seed builds the workbook (sizes, places, how many tables, links,
notes, names and formats, an insert before or after the object is held) and samples the calls, every class and member at least
once, each late bound and typed, on three kinds of book: built new, built and saved and reopened by the same engine, and two books
Excel saved. Answers carry their TypeName.

    python scripts/measure_table_pivot_rounds.py predict <seed>
    python scripts/measure_table_pivot_rounds.py measure <seed>

``predict`` writes the model's answers and their digest to tests/fixtures/table_pivot_rounds/round_<seed>.json before Excel is
asked anything; post the digest on the coordination board. ``measure`` asks live Excel in short runs, each under
excel-coordination/live.lock, adds its answers, and prints every difference. tests/test_excel_table_pivot_rounds.py replays the
rounds. A round the model has been fixed against is used up: the next test needs a new seed.
"""
from __future__ import annotations

import hashlib
import json
import os
import random
import shutil
import subprocess
import sys
import tempfile
import time
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
FIXTURES = ROOT / 'tests' / 'fixtures'
OUT = FIXTURES / 'table_pivot_rounds'
#: Excel saved both: a table and a pivot table; links, notes, names and conditional formats.
TABLE_BOOK = FIXTURES / 'dxf_table' / 'made.xlsx'
LINK_BOOK = FIXTURES / 'object_roundtrip.xlsx'
KINDS = ('new', 'reopened', 'file')
#: How many calls a round samples, each asked late bound and typed on one kind of book.
SAMPLE = 150
CHUNK = 25

HELPER = '''
Private Function Key(ByVal v As Variant, ByVal book As String) As String
    Dim t As String, n As String
    On Error Resume Next
    If IsObject(v) Then
        If v Is Nothing Then Key = "Nothing": Exit Function
        t = TypeName(v)
        Select Case t
        Case "Range": n = v.Address(False, False) & "@" & v.Worksheet.Name
        Case "Application": n = CStr(v Is Application)
        Case "Workbook": n = "<book>"
        Case Else
            n = v.Name
            If Err.Number <> 0 Then Err.Clear: n = "?"
        End Select
        Key = t & ":" & Replace(n, book, "<book>")
    ElseIf IsNull(v) Then
        Key = "Null"
    ElseIf IsArray(v) Then
        Key = "Array"
    Else
        Key = TypeName(v) & ":" & Replace(CStr(v), book, "<book>")
    End If
End Function
'''

#: Each class: its typed declaration and the ways a macro reaches it, from the variables the setup leaves.
REACHES: dict[str, list[str]] = {
    'ListObject': ['ws.ListObjects(1)', 'ws.ListObjects(tname)', 'ws.Range(tcell).ListObject', 'ws.ListObjects(ws.ListObjects.Count)'],
    'ListObjects': ['ws.ListObjects'],
    'ListColumn': ['ws.ListObjects(1).ListColumns(1)', 'ws.ListObjects(1).ListColumns(tcolumn)', 'ws.ListObjects(1).ListColumns(ws.ListObjects(1).ListColumns.Count)'],
    'ListColumns': ['ws.ListObjects(1).ListColumns'],
    'ListRow': ['ws.ListObjects(1).ListRows(1)', 'ws.ListObjects(1).ListRows(ws.ListObjects(1).ListRows.Count)'],
    'ListRows': ['ws.ListObjects(1).ListRows'],
    'PivotTable': ['ws2.PivotTables(1)', 'ws2.PivotTables(ptname)', 'ws2.Range(ptcell).PivotTable'],
    'PivotTables': ['ws2.PivotTables'],
    'PivotField': ['ws2.PivotTables(1).PivotFields(pfield)', 'ws2.PivotTables(1).PivotFields(1)', 'ws2.PivotTables(1).DataFields(1)'],
    'PivotFields': ['ws2.PivotTables(1).PivotFields', 'ws2.PivotTables(1).RowFields'],
    'PivotItem': ['ws2.PivotTables(1).PivotFields(pfield).PivotItems(1)', 'ws2.PivotTables(1).PivotFields(pfield).PivotItems(pitem)'],
    'PivotItems': ['ws2.PivotTables(1).PivotFields(pfield).PivotItems'],
    'PivotCache': ['ws2.PivotTables(1).PivotCache', 'wbT.PivotCaches(1)'],
    'PivotCaches': ['wbT.PivotCaches'],
    'Hyperlink': ['wsL.Hyperlinks(1)', 'wsL.Range(hcell).Hyperlinks(1)'],
    'Hyperlinks': ['wsL.Hyperlinks', 'wsL.Range(hcell).Hyperlinks'],
    'Comment': ['wsL.Comments(1)', 'wsL.Range(ccell).Comment', 'wsL.Comments(wsL.Comments.Count)'],
    'Comments': ['wsL.Comments'],
    'Name': ['wbL.Names(1)', 'wbL.Names(nname)', 'wbL.Names(wbL.Names.Count)'],
    'Names': ['wbL.Names'],
    'FormatCondition': ['wsL.Range(fcells).FormatConditions(1)'],
    'FormatConditions': ['wsL.Range(fcells).FormatConditions', 'wsL.Cells.FormatConditions'],
}
TRIO = ['x = Key(o.Application, book)', 'x = Key(o.Creator, book)', 'x = Key(o.Parent, book)']
#: Late bound only: assigning to a member the type library holds read-only does not compile typed.
ASSIGNS = ['o.Creator = 0', 'Set o.Application = Nothing', 'Set o.Parent = Nothing']
MEMBERS: dict[str, list[str]] = {
    'ListObject': ['x = Key(o.Range, book)', 'x = Key(o.DataBodyRange, book)', 'x = Key(o.HeaderRowRange, book)', 'x = Key(o.TotalsRowRange, book)',
                   'x = Key(o.ListColumns, book) & o.ListColumns.Count', 'x = Key(o.ListRows, book) & o.ListRows.Count', 'o.Delete: x = ws.ListObjects.Count'],
    'ListColumn': ['x = Key(o.Range, book)', 'x = Key(o.DataBodyRange, book)', 'x = Key(o.Index, book)'],
    'ListRow': ['x = Key(o.Range, book)', 'x = Key(o.Index, book)'],
    'ListColumns': ['x = Key(o.Add, book) & o.Count', 'x = Key(o.Add(1), book) & o.Count', 'x = Key(o.Add(o.Count + 1), book) & o.Count', 'x = Key(o.Add(o.Count + 2), book) & o.Count'],
    'ListRows': ['x = Key(o.Add, book) & o.Count', 'x = Key(o.Add(1), book) & o.Count', 'x = Key(o.Add(o.Count + 1), book) & o.Count', 'x = Key(o.Add(0), book) & o.Count'],
    'Hyperlink': ['x = Key(o.Range, book)', 'x = Key(o.Name, book)', 'o.Delete: x = wsL.Hyperlinks.Count'],
    'Hyperlinks': ['o.Delete: x = wsL.Hyperlinks.Count'],
    'Comment': ['x = Key(o.Author, book)', 'x = Key(o.Text, book)', 'x = Key(o.Next, book)', 'x = Key(o.Previous, book)', 'x = Key(o.Text("New"), book) & o.Text',
                'x = Key(o.Text("Add", 2), book) & o.Text', 'x = Key(o.Text("Over", 1, True), book) & o.Text', 'o.Delete: x = wsL.Comments.Count'],
}
COLLECTIONS = ('ListObjects', 'ListColumns', 'ListRows', 'PivotTables', 'PivotFields', 'PivotItems', 'PivotCaches', 'Hyperlinks', 'Comments', 'Names', 'FormatConditions')
#: What an Item index can be; the names come from the setup's variables.
INDEXES = ['1', '0', '-1', 'o.Count', 'o.Count + 1', '1.6', '"1"', 'CLng(2)', 'True', 'Empty', '"nosuch"']
NAMED = {'ListObjects': 'tname', 'ListColumns': 'tcolumn', 'PivotTables': 'ptname', 'PivotFields': 'pfield', 'PivotItems': 'pitem', 'Names': 'nname'}


def setup(kind: str, seed: int) -> str:
    """VBA that leaves wbT, ws, ws2 (table and pivot), wbL, wsL (links, notes, names, formats) and the names the reaches use."""
    rng = random.Random(f'{seed}:{kind}')
    if kind == 'file':
        return ('    Set wbT = Workbooks.Open(path & "\\t.xlsx")\n    Set ws = wbT.Worksheets("one"): Set ws2 = wbT.Worksheets("pivot")\n'
                '    Set wbL = Workbooks.Open(path & "\\l.xlsx")\n    Set wsL = wbL.Worksheets("Data")\n'
                '    tname = ws.ListObjects(1).Name: tcell = "E3": tcolumn = ws.ListObjects(1).ListColumns(1).Name\n'
                '    ptname = ws2.PivotTables(1).Name: ptcell = "A4": pfield = "Name": pitem = ws2.PivotTables(1).PivotFields("Name").PivotItems(1).Name\n'
                '    hcell = "B10": ccell = "A11": nname = "Constant": fcells = "D10:D12"\n')
    rows = rng.randint(2, 6)
    column = rng.randint(5, 9)
    table_rows = rng.randint(1, 5)
    keys = [rng.choice('abcd') for _ in range(rows)]
    lines = ['    Set wbT = Workbooks.Add(xlWBATWorksheet): Set wbL = wbT', '    Set ws = wbT.Worksheets(1): ws.Name = "Main": Set wsL = ws',
             '    Set ws2 = wbT.Worksheets.Add(After:=ws): ws2.Name = "Second": ws.Activate', '    ws.Range("A1:B1").Value = Array("k", "v")']
    lines += [f'    ws.Range("A{at + 2}:B{at + 2}").Value = Array("{key}", {rng.randint(1, 50)})' for at, key in enumerate(keys)]
    letter = chr(64 + column)
    after = chr(65 + column)
    lines += [f'    ws.Range("{letter}1:{after}1").Value = Array("h{seed % 7}", "g")', f'    ws.Range("{letter}2:{after}{table_rows + 1}").Value = {rng.randint(1, 9)}',
              f'    Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("{letter}1:{after}{table_rows + 1}"), , xlYes): lo.Name = "Tb{seed % 97}"']
    if rng.random() < 0.4:
        lines.append(f'    ws.Range("M1:M3").Value = 1: ws.ListObjects.Add(xlSrcRange, ws.Range("M1:M3"), , xlYes).Name = "Tz{seed % 89}"')
    if rng.random() < 0.4:
        lines.append('    lo.ShowTotals = True')
    lines += [f'    Set pt = wbT.PivotCaches.Create(xlDatabase, ws.Range("A1:B{rows + 1}")).CreatePivotTable(ws2.Range("A3"), "PT{seed % 83}")',
              '    pt.PivotFields("k").Orientation = xlRowField', '    pt.AddDataField pt.PivotFields("v"), "Sum of v", xlSum']
    links = rng.randint(1, 3)
    lines += [f'    ws.Hyperlinks.Add ws.Range("Q{at + 2}"), "http://example.com/{seed}/{at}", "", "", "L{at}"' for at in range(links)]
    if rng.random() < 0.5:
        lines.append('    ws.Hyperlinks.Add ws.Range("R2"), "", "Main!A1", "", "Inner"')
    notes = rng.sample(range(2, 12), rng.randint(1, 3))
    lines += [f'    ws.Range("S{row}").AddComment "n{row}"' for row in notes]
    names = rng.randint(1, 3)
    lines += [f'    wbT.Names.Add "Nm_{at}_{seed % 10}", "=Main!$A${at + 1}"' for at in range(names)]
    if rng.random() < 0.5:
        lines.append('    ws.Names.Add "Local1", "=Main!$B$2"')
    formats = rng.randint(1, 3)
    operators = rng.sample([3, 4, 5, 6, 7, 8], formats)
    lines += [f'    ws.Range("D1:D5").FormatConditions.Add xlCellValue, {op}, "={rng.randint(1, 5)}"' for op in operators]
    if rng.random() < 0.4:
        lines.append('    ws.Range("D1:D5").FormatConditions.Add xlExpression, , "=D1>2"')
    first_key = sorted(set(keys))[0]
    lines += [f'    tname = lo.Name: tcell = "{letter}2": tcolumn = "g"', f'    ptname = pt.Name: ptcell = "A4": pfield = "k": pitem = "{first_key}"',
              f'    hcell = "Q2": ccell = "S{notes[0]}": nname = "Nm_0_{seed % 10}": fcells = "D1:D5"']
    if kind == 'reopened':
        lines += ['    wbT.SaveAs path & "\\r.xlsx", xlOpenXMLWorkbook', '    wbT.Close False', '    Set wbT = Workbooks.Open(path & "\\r.xlsx"): Set wbL = wbT',
                  '    Set ws = wbT.Worksheets("Main"): Set wsL = ws: Set ws2 = wbT.Worksheets("Second")']
    return '\n'.join(lines) + '\n'


def calls(seed: int) -> list[dict[str, str]]:
    """Every class and member at least once, then a seeded fill, each with its kind of book, its reach and whether the object is
    held across an insert that moves it."""
    rng = random.Random(seed)
    every: list[dict[str, str]] = []
    for cls, reaches in REACHES.items():
        statements = [*TRIO, *ASSIGNS, *MEMBERS.get(cls, [])]
        if cls in COLLECTIONS:
            statements += ['x = Key(o.Count, book)'] + [f'x = Key(o.Item({index}), book)' for index in INDEXES]
            if cls in NAMED:
                statements += [f'x = Key(o.Item({NAMED[cls]}), book)', f'x = Key(o.Item(UCase({NAMED[cls]})), book)']
        for statement in statements:
            every += [{'cls': cls, 'reach': reach, 'call': statement} for reach in reaches]
    by_member: dict[tuple[str, str], list[dict[str, str]]] = {}
    for one in every:
        member = one['call'].split('(o.')[-1].split('.')[-1] if '(o.' in one['call'] else one['call']
        by_member.setdefault((one['cls'], member.split('(')[0].split(' ')[0].split(',')[0]), []).append(one)
    chosen = [rng.choice(group) for group in by_member.values()]
    rest = [one for one in every if one not in chosen]
    chosen += rng.sample(rest, max(0, SAMPLE - len(chosen)))
    edits = ['', 'ws.Rows(1).Insert', 'ws.Columns(1).Insert', 'wsL.Rows(1).Insert']
    for one in chosen:
        one['kind'] = rng.choice(KINDS)
        one['edit'] = rng.choice(edits)
        one['held'] = str(rng.random() < 0.4)
    return chosen


def source(case: dict[str, str], typed: bool, seed: int) -> str:
    """A function One(path) that builds the book, runs the call and closes what it opened."""
    kind = case['cls'] if typed else 'Object'
    edit = case['edit']
    hold = f'    Set o = {case["reach"]}\n'
    body = (hold + (f'    {edit}\n' if edit else '')) if case['held'] == 'True' else ((f'    {edit}\n' if edit else '') + hold)
    statement = case['call']
    if typed and statement in ASSIGNS:
        statement = 'x = "late only"'
    return (HELPER + 'Public Function One(ByVal path As String) As String\n'
            f'    Dim wbT As Object, wbL As Object, ws As Object, ws2 As Object, wsL As Object, lo As Object, pt As Object, o As {kind}, x As String, book As String\n'
            '    Dim tname As String, tcell As String, tcolumn As String, ptname As String, ptcell As String, pfield As String, pitem As String\n'
            '    Dim hcell As String, ccell As String, nname As String, fcells As String\n'
            '    Application.DisplayAlerts = False\n' + setup(case['kind'], seed) + '    book = wbT.Name\n    On Error Resume Next\n' + body
            + '    If Err.Number <> 0 Then\n        One = "reach E" & Err.Number\n    Else\n'
            f'        {statement}\n        One = x & " E" & Err.Number\n    End If\n    Err.Clear\n    On Error GoTo 0\n'
            '    If Not wbL Is wbT Then wbL.Close False\n    wbT.Close False\nEnd Function\n')


def books(folder: Path) -> None:
    shutil.copyfile(TABLE_BOOK, folder / 't.xlsx')
    shutil.copyfile(LINK_BOOK, folder / 'l.xlsx')


def model_answer(case: dict[str, str], typed: bool, seed: int) -> str:
    from pyopenvba.apps.excel import ExcelApplication
    from pyopenvba.exceptions import VBAError

    with tempfile.TemporaryDirectory() as held:
        folder = Path(held)
        books(folder)
        app = ExcelApplication(sandbox=False)
        app.add_workbook()
        app.add_module(source(case, typed, seed), name='Round')
        try:
            return str(app.run('Round.One', str(folder)))
        except VBAError as failed:
            return f'model {type(failed).__name__}: {str(failed)[:120]}'


def digest(cases: list[dict[str, str]]) -> str:
    text = json.dumps([[case['cls'], case['reach'], case['call'], case['kind'], case['edit'], case['held'], case['model_late'], case['model_typed']] for case in cases])
    return hashlib.sha256(text.encode('utf-8')).hexdigest()[:16]


def predict(seed: int) -> None:
    path = OUT / f'round_{seed}.json'
    assert not path.exists(), f'{path} exists: a round is predicted once'
    cases = calls(seed)
    for case in cases:
        case['model_late'] = model_answer(case, False, seed)
        case['model_typed'] = model_answer(case, True, seed)
    OUT.mkdir(parents=True, exist_ok=True)
    record = {'seed': seed, 'predicted_at': datetime.now(timezone.utc).isoformat(), 'predictions_sha256_16': digest(cases), 'cases': cases}
    path.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    print('predicted', len(cases), 'cases, digest', record['predictions_sha256_16'])


def _lock() -> Path:
    common = subprocess.run(['git', 'rev-parse', '--git-common-dir'], capture_output=True, text=True, check=True, cwd=ROOT).stdout.strip()
    return (ROOT / common).resolve() / 'excel-coordination' / 'live.lock'


def measure(seed: int) -> None:
    from pyvbaharness import ExcelSession, HarnessConfig

    path = OUT / f'round_{seed}.json'
    record = json.loads(path.read_text(encoding='utf-8'))
    cases = record['cases']
    assert digest(cases) == record['predictions_sha256_16'], 'the predictions changed after they were written'
    lock = _lock()
    with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as held:
        folder = Path(held)
        books(folder)
        for tag, typed in (('late', False), ('typed', True)):
            todo = [case for case in cases if tag not in case]
            for start in range(0, len(todo), CHUNK):
                chunk = todo[start:start + CHUNK]
                while True:
                    try:
                        os.mkdir(lock)
                        break
                    except FileExistsError:
                        time.sleep(10)
                try:
                    with ExcelSession(HarnessConfig(lock_wait_s=1200.0, exclusive=False)) as excel:
                        excel.new_document()
                        for case in chunk:
                            done = excel.run_vba(source(case, typed, seed), 'One', (str(folder),), timeout=60.0, line_numbers=False)
                            case[tag] = str(done.value) if done.ok else f'no answer: {str(done)[:120]}'
                finally:
                    os.rmdir(lock)
                path.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
                print(tag, start + len(chunk), 'of', len(todo), flush=True)
    record['measured_at'] = datetime.now(timezone.utc).isoformat()
    path.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    differ = [(case, tag) for case in cases for tag in ('late', 'typed') if case[tag] != case[f'model_{tag}']]
    for case, tag in differ:
        print(tag, case['kind'], case['cls'], case['reach'], '|', case['call'], '| held', case['held'], case['edit'], '| Excel', case[tag], '| model', case[f'model_{tag}'])
    print('differ', len(differ), 'of', 2 * len(cases))


def main() -> None:
    action, seed = sys.argv[1], int(sys.argv[2])
    predict(seed) if action == 'predict' else measure(seed)


if __name__ == '__main__':
    main()
