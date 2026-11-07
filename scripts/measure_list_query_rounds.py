"""Held-out rounds over ListObject.QueryTable on a table over a database, and the members a macro reads through it, on the table rounds'
engine (scripts/measure_table_pivot_rounds.py): seeded calls, every member at least once, each late bound and typed, on a book built
new, one built, saved and reopened, and Excel's own book (tests/fixtures/query_refresh_loaded/access_table.xlsx, its folder written as
the run's). The table is ListObjects.Add over an Access file the round's own VBA makes through Excel's ACE OLEDB provider in the
run's folder; the model's Refresh function stands for it. No network source is ever named, and nothing here touches encryption,
signing or a certificate.

Left out: the connection string, which Excel reads back fuller after a refresh (a named difference), and refreshes that change the
table's columns, which belong to QueryTable.Refresh's own row.

    python scripts/measure_list_query_rounds.py predict <seed>
    python scripts/measure_list_query_rounds.py measure <seed>

tests/test_excel_list_query_rounds.py replays the rounds.
"""
from __future__ import annotations

import json
import random
import re
import sys
import tempfile
from datetime import datetime, timezone
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import measure_table_pivot_rounds as base  # noqa: E402
from measure_query_refresh import PROVIDER  # noqa: E402

OUT = base.FIXTURES / 'list_query_rounds'
#: Excel's own book with a ListObject over the Access file, its folder written C:\\data.
FILE_BOOK = base.FIXTURES / 'query_refresh_loaded' / 'access_table.xlsx'
REACHES = {'ListObject': ['ws.ListObjects(1)', 'ws.ListObjects(tname)', 'ws.Range(tcell).ListObject'],
           'QueryTable': ['ws.ListObjects(1).QueryTable', 'ws.ListObjects(tname).QueryTable', 'ws.Range(tcell).ListObject.QueryTable']}
_SETTINGS = ('CommandText', 'CommandType', 'QueryType', 'BackgroundQuery', 'RefreshStyle', 'PreserveColumnInfo', 'FieldNames', 'Name',
             'RowNumbers', 'FillAdjacentFormulas', 'PreserveFormatting', 'AdjustColumnWidth', 'SavePassword', 'SaveData', 'EnableRefresh',
             'RefreshOnFileOpen', 'EnableEditing', 'MaintainConnection', 'RefreshPeriod', 'Refreshing')
MEMBERS = {
    'ListObject': ['x = Key(o.QueryTable, book)', 'x = TypeName(o.QueryTable) & " " & Key(o.QueryTable.ListObject, book)', 'x = Key(o.SourceType, book)',
                   'x = Key(o.Range, book)', 'o.Refresh: x = Key(o.Range, book)', 'x = Key(o.QueryTable.ResultRange, book)',
                   'o.Unlist: x = ws.ListObjects.Count & " " & wbT.Connections.Count & " " & ws.QueryTables.Count',
                   'o.Delete: x = ws.ListObjects.Count & " " & wbT.Connections.Count', 'x = (o.QueryTable Is o.QueryTable)'],
    'QueryTable': [*[f'x = Key(o.{name}, book)' for name in _SETTINGS],
                   'x = Key(o.ResultRange, book)', 'x = Key(o.Destination, book)', 'x = Key(o.ListObject, book)', 'x = Key(o.WorkbookConnection, book)',
                   'o.CommandText = "SELECT * FROM T WHERE Id <= 4 ORDER BY Id": o.Refresh False: x = Key(o.ResultRange, book)',
                   'o.CommandText = "SELECT * FROM T WHERE Id <= 1 ORDER BY Id": o.Refresh False: x = Key(o.ResultRange, book) & " " & Key(o.Refreshing, book)',
                   'o.CommandType = xlCmdTable: o.CommandText = "T": o.Refresh False: x = Key(o.ResultRange, book)',
                   'o.BackgroundQuery = True: x = Key(o.BackgroundQuery, book)', 'o.RefreshStyle = xlOverwriteCells: x = Key(o.RefreshStyle, book)',
                   'o.SaveData = False: x = Key(o.SaveData, book)', 'o.Refresh: x = Key(o.ResultRange, book)'],
}
_DATABASE = ('    If Dir(path & "\\rows.accdb") = "" Then\n'
             '        With CreateObject("ADOX.Catalog")\n'
             f'            .Create "Provider={PROVIDER};Data Source=" & path & "\\rows.accdb"\n'
             '            .ActiveConnection.Execute "CREATE TABLE T (Id LONG, Name TEXT(10))"\n'
             + ''.join(f'            .ActiveConnection.Execute "INSERT INTO T VALUES ({at}, \'n{at}\')"\n' for at in range(1, 7))
             + '            .ActiveConnection.Close\n'
             '        End With\n'
             '    End If\n')


def setup(kind: str, seed: int) -> str:
    """VBA that leaves wbT, ws (the table's sheet), ws2, wbL and wsL, and tname and tcell for the table."""
    if kind == 'file':
        return ('    Set wbT = Workbooks.Open(path & "\\f.xlsx"): Set wbL = wbT\n'
                '    Set ws = wbT.Worksheets(1): Set wsL = ws: Set ws2 = ws\n'
                '    tname = ws.ListObjects(1).Name: tcell = "C3"\n')
    rng = random.Random(f'{seed}:{kind}:list')
    row, column, rows = rng.randint(2, 5), rng.randint(2, 5), rng.randint(1, 5)
    corner = f'{chr(64 + column)}{row}'
    refreshed = kind == 'reopened' or rng.random() < 0.7
    lines = [_DATABASE.rstrip('\n'),
             '    Set wbT = Workbooks.Add(xlWBATWorksheet): Set wbL = wbT',
             '    Set ws = wbT.Worksheets(1): ws.Name = "Main": Set wsL = ws',
             '    Set ws2 = wbT.Worksheets.Add(After:=ws): ws2.Name = "Second": ws.Activate',
             f'    Set lo = ws.ListObjects.Add(xlSrcExternal, Array("OLEDB;Provider={PROVIDER};Data Source=" & path & "\\rows.accdb"), , xlYes, ws.Range("{corner}"))']
    if refreshed:
        if rng.random() < 0.2:
            lines.append('    lo.QueryTable.CommandType = xlCmdTable: lo.QueryTable.CommandText = Array("T")')
        else:
            lines.append(f'    lo.QueryTable.CommandText = "SELECT * FROM T WHERE Id <= {rows} ORDER BY Id"')
        lines.append('    lo.QueryTable.Refresh False')
    if rng.random() < 0.3:
        lines.append(f'    ws.Range("{chr(64 + column)}{row + 9}").Value = "below"')
    lines.append(f'    tname = lo.Name: tcell = "{corner}"')
    if kind == 'reopened':
        lines += ['    saved = path & "\\r" & Format(Now, "hhnnss") & (Int(Timer * 1000) Mod 1000) & ".xlsx"', '    wbT.SaveAs saved, xlOpenXMLWorkbook',
                  '    wbT.Close False', '    Set wbT = Workbooks.Open(saved): Set wbL = wbT',
                  '    Set ws = wbT.Worksheets("Main"): Set wsL = ws: Set ws2 = wbT.Worksheets("Second")']
    return '\n'.join(lines) + '\n'


def books(folder: Path) -> None:
    """Excel's own book, its connection pointed at the run's folder."""
    import io
    import zipfile

    found = io.BytesIO()
    with zipfile.ZipFile(FILE_BOOK) as source, zipfile.ZipFile(found, 'w', zipfile.ZIP_DEFLATED) as target:
        for info in source.infolist():
            target.writestr(info, source.read(info).replace(b'C:\\data', str(folder).encode('utf-8')))
    (folder / 'f.xlsx').write_bytes(found.getvalue())


LISTS = base.Plan(REACHES, MEMBERS, (), {}, setup, books, base.DIMS, OUT, base.SAMPLE)


def table_t(request: object) -> list[list[object]]:
    """The rows of T the command asks for, as ACE gives them: a table's name gives the whole table, and SQL text runs as SQL whatever
    CommandType says (tests/fixtures/list_query_add.json, table_then_sql)."""
    source = getattr(request, 'source')
    if source.command_text == 'T':
        last = 6
    else:
        found = re.fullmatch(r'SELECT \* FROM T WHERE Id <= (\d+) ORDER BY Id', source.command_text)
        assert found is not None, source.command_text
        last = int(found.group(1))
    return [['Id', 'Name'], *[[float(at), f'n{at}'] for at in range(1, last + 1)]]


def model_answer(case: dict[str, str], typed: bool, seed: int) -> str:
    from pyopenvba.apps.excel import ExcelApplication
    from pyopenvba.exceptions import VBAError

    with tempfile.TemporaryDirectory() as held:
        folder = Path(held)
        books(folder)
        # The model's Refresh function stands for the file, so the round's VBA finds one and makes none.
        (folder / 'rows.accdb').write_bytes(b'')
        app = ExcelApplication(sandbox=False)
        app.add_workbook()
        app.set_hook('Refresh', table_t)
        app.add_module(base.source(case, typed, seed, LISTS), name='Round')
        try:
            return str(app.run('Round.One', str(folder)))
        except VBAError as failed:
            return f'model {type(failed).__name__}: {str(failed)[:120]}'


def predict(seed: int) -> None:
    path = OUT / f'round_{seed}.json'
    assert not path.exists(), f'{path} exists: a round is predicted once'
    cases = base.calls(seed, LISTS)
    for case in cases:
        case['model_late'] = model_answer(case, False, seed)
        case['model_typed'] = model_answer(case, True, seed)
    OUT.mkdir(parents=True, exist_ok=True)
    record = {'seed': seed, 'predicted_at': datetime.now(timezone.utc).isoformat(), 'predictions_sha256_16': base.digest(cases), 'cases': cases}
    path.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    print('predicted', len(cases), 'cases, digest', record['predictions_sha256_16'])


if __name__ == '__main__':
    action, number = sys.argv[1], int(sys.argv[2])
    predict(number) if action == 'predict' else base.measure(number, int(sys.argv[3]) if len(sys.argv) > 3 else 1000, LISTS)
