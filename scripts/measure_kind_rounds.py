"""Held-out rounds over the VERIFIED rows the table and pivot rounds do not reach: each kind of conditional format rule, colour
scale criteria, icon sets, a sheet's and a table's AutoFilter and its Filters, pivot cells, their row items, pivot filters, slicers
and threaded comments. The machinery is scripts/measure_table_pivot_rounds.py's; this script gives it its own book and reaches.

    python scripts/measure_kind_rounds.py book
    python scripts/measure_kind_rounds.py predict <seed>
    python scripts/measure_kind_rounds.py measure <seed>

``book`` has Excel build and save tests/fixtures/kind_rounds/k.xlsx, the book the 'file' kind opens; it holds no threaded comment,
whose person list would carry the measuring user's account, so that kind adds its thread after opening. ``predict`` and
``measure`` are as in the table rounds; tests/test_excel_kind_rounds.py replays the rounds.
"""
from __future__ import annotations

import os
import random
import shutil
import sys
import tempfile
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import measure_table_pivot_rounds as base  # noqa: E402

OUT = base.FIXTURES / 'kind_rounds'
BOOK = OUT / 'k.xlsx'
#: The seed the saved book was built from, so the 'file' kind knows where everything is.
BOOK_SEED = 52100

REACHES: dict[str, list[str]] = {
    'Top10': ['ws.Range(rTop).FormatConditions(1)'],
    'AboveAverage': ['ws.Range(rAbove).FormatConditions(1)'],
    'UniqueValues': ['ws.Range(rUnique).FormatConditions(1)'],
    'ColorScale': ['ws.Range(rScale).FormatConditions(1)'],
    'ColorScaleCriteria': ['ws.Range(rScale).FormatConditions(1).ColorScaleCriteria'],
    'ColorScaleCriterion': ['ws.Range(rScale).FormatConditions(1).ColorScaleCriteria(1)',
                            'ws.Range(rScale).FormatConditions(1).ColorScaleCriteria(ws.Range(rScale).FormatConditions(1).ColorScaleCriteria.Count)'],
    'Databar': ['ws.Range(rBar).FormatConditions(1)'],
    'IconSetCondition': ['ws.Range(rIcon).FormatConditions(1)'],
    'IconSet': ['ws.Range(rIcon).FormatConditions(1).IconSet', 'wbT.IconSets(iset)'],
    'IconSets': ['wbT.IconSets'],
    'AutoFilter': ['ws.AutoFilter', 'ws.ListObjects(1).AutoFilter'],
    'Filters': ['ws.AutoFilter.Filters', 'ws.ListObjects(1).AutoFilter.Filters'],
    'Filter': ['ws.AutoFilter.Filters(1)', 'ws.AutoFilter.Filters(ws.AutoFilter.Filters.Count)', 'ws.ListObjects(1).AutoFilter.Filters(1)'],
    'PivotCell': ['ws2.PivotTables(1).TableRange1.Cells(2, 1).PivotCell', 'ws2.PivotTables(1).TableRange1.Cells(2, 2).PivotCell'],
    'PivotItemList': ['ws2.PivotTables(1).TableRange1.Cells(2, 2).PivotCell.RowItems'],
    'PivotFilters': ['ws2.PivotTables(1).PivotFields(pfield).PivotFilters', 'ws2.PivotTables(1).ActiveFilters'],
    'Slicers': ['ws2.PivotTables(1).Slicers', 'ws.ListObjects(1).Slicers'],
    'CommentThreaded': ['ws.Range(tcell).CommentThreaded', 'ws.CommentsThreaded(1)'],
    'CommentsThreaded': ['ws.CommentsThreaded', 'ws.Range(tcell).CommentThreaded.Replies'],
}
#: Count and Item only where their rows are VERIFIED, or PARTIAL on an index the rounds never pass (Null).
COLLECTIONS = ('Filters', 'IconSet', 'IconSets', 'PivotItemList')
DIMS = ('    Dim wbT As Object, wbL As Object, ws As Object, ws2 As Object, wsL As Object, lo As Object, pt As Object, o As {kind}, x As String, book As String\n'
        '    Dim rTop As String, rAbove As String, rUnique As String, rScale As String, rBar As String, rIcon As String\n'
        '    Dim pfield As String, tcell As String, saved As String, iset As Long\n')


def _build(seed: int) -> tuple[list[str], list[str]]:
    """The VBA that builds the book on a new workbook, and the lines that set the reaches' variables."""
    rng = random.Random(f'{seed}:build')
    rows = rng.randint(3, 6)
    keys = [rng.choice('abcd') for _ in range(rows)]
    lines = ['    Set ws = wbT.Worksheets(1): ws.Name = "Main"', '    Set ws2 = wbT.Worksheets.Add(After:=ws): ws2.Name = "Second": ws.Activate',
             '    ws.Range("A1:B1").Value = Array("k", "v")']
    lines += [f'    ws.Range("A{at + 2}:B{at + 2}").Value = Array("{key}", {rng.randint(1, 50)})' for at, key in enumerate(keys)]
    lines += [f'    Set pt = wbT.PivotCaches.Create(xlDatabase, ws.Range("A1:B{rows + 1}")).CreatePivotTable(ws2.Range("A3"), "PK{seed % 83}")',
              '    pt.PivotFields("k").Orientation = xlRowField', '    pt.AddDataField pt.PivotFields("v"), "Sum of v", xlSum']
    if rng.random() < 0.5:
        lines.append(f'    pt.PivotFields("k").PivotFilters.Add2 xlCaptionEquals, , "{keys[0]}"')
    first = rng.randint(4, 8)
    height = rng.randint(3, 6)
    spans = [f'{chr(64 + first + at)}1:{chr(64 + first + at)}{height}' for at in range(6)]
    lines += [f'    ws.Range("{span}").Value = {rng.randint(1, 9)}: ws.Range("{span.split(":")[0]}").Value = {rng.randint(10, 20)}' for span in spans]
    adds = ['AddTop10', 'AddAboveAverage', 'AddUniqueValues', f'AddColorScale({rng.choice((2, 3))})', 'AddDatabar', 'AddIconSetCondition']
    lines += [f'    ws.Range("{span}").FormatConditions.{add}' for span, add in zip(spans, adds)]
    if rng.random() < 0.5:
        lines.append(f'    ws.Range("{spans[5]}").FormatConditions(1).IconSet = wbT.IconSets({rng.randint(1, 20)})')
    table = chr(64 + first + 7)
    after = chr(64 + first + 8)
    lines += [f'    ws.Range("{table}1:{after}1").Value = Array("p", "q"): ws.Range("{table}2:{after}{rng.randint(2, 5)}").Value = {rng.randint(1, 9)}',
              f'    Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("{table}1:{after}5"), , xlYes): lo.Name = "Tk{seed % 97}"']
    if rng.random() < 0.5:
        lines.append('    lo.Range.AutoFilter 1, "<5"')
    lines.append(f'    ws.Range("A1:B{rows + 1}").AutoFilter')
    if rng.random() < 0.6:
        lines.append(f'    ws.Range("A1:B{rows + 1}").AutoFilter {rng.randint(1, 2)}, "{rng.choice(("a", ">3", "<>b"))}"')
    names = dict(zip(('rTop', 'rAbove', 'rUnique', 'rScale', 'rBar', 'rIcon'), spans))
    values = [f'    {name} = "{span}"' for name, span in names.items()]
    values.append(f'    pfield = "k": iset = {rng.randint(1, 20)}: tcell = "{chr(64 + first + 10)}{rng.randint(2, 6)}"')
    return lines, values


def _thread(seed: int) -> list[str]:
    rng = random.Random(f'{seed}:thread')
    lines = [f'    ws.Range(tcell).AddCommentThreaded "t{seed % 10}"']
    if rng.random() < 0.5:
        lines.append('    ws.Range(tcell).CommentThreaded.AddReply "r"')
    if rng.random() < 0.4:
        lines.append(f'    ws.Range(tcell).Offset({rng.randint(1, 3)}, 0).AddCommentThreaded "u"')
    return lines


def setup(kind: str, seed: int) -> str:
    """VBA that leaves wbT, ws, ws2, lo, pt, wsL = ws, wbL = wbT and the reaches' variables."""
    if kind == 'file':
        _, values = _build(BOOK_SEED)
        lines = ['    Set wbT = Workbooks.Open(path & "\\k.xlsx"): Set wbL = wbT', '    Set ws = wbT.Worksheets("Main"): Set wsL = ws',
                 '    Set ws2 = wbT.Worksheets("Second")', *values, *_thread(seed)]
        return '\n'.join(lines) + '\n'
    lines, values = _build(seed)
    lines = ['    Set wbT = Workbooks.Add(xlWBATWorksheet): Set wbL = wbT', *lines, '    Set wsL = ws', *values, *_thread(seed)]
    if kind == 'reopened':
        lines += ['    saved = path & "\\r" & Format(Now, "hhnnss") & (Int(Timer * 1000) Mod 1000) & ".xlsx"', '    wbT.SaveAs saved, xlOpenXMLWorkbook',
                  '    wbT.Close False', '    Set wbT = Workbooks.Open(saved): Set wbL = wbT',
                  '    Set ws = wbT.Worksheets("Main"): Set wsL = ws: Set ws2 = wbT.Worksheets("Second")']
    return '\n'.join(lines) + '\n'


def books(folder: Path) -> None:
    shutil.copyfile(BOOK, folder / 'k.xlsx')


KINDS = base.Plan(REACHES, {}, COLLECTIONS, {}, setup, books, DIMS, OUT, base.SAMPLE)


def make_book() -> None:
    """Excel builds the book from BOOK_SEED, without a thread, and saves it."""
    from pyvbaharness import ExcelSession, HarnessConfig

    lines, _ = _build(BOOK_SEED)
    source = ('Public Function Make(ByVal target As String) As String\n'
              '    Dim wbT As Object, ws As Object, ws2 As Object, lo As Object, pt As Object\n    Application.DisplayAlerts = False\n'
              '    Set wbT = Workbooks.Add(xlWBATWorksheet)\n' + '\n'.join(lines) + '\n'
              '    wbT.SaveAs target, xlOpenXMLWorkbook\n    wbT.Close False\n    Make = "saved"\nEnd Function\n')
    OUT.mkdir(parents=True, exist_ok=True)
    lock = base._lock()
    while True:
        try:
            os.mkdir(lock)
            break
        except FileExistsError:
            time.sleep(10)
    try:
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as held, ExcelSession(HarnessConfig(lock_wait_s=1200.0, exclusive=False)) as excel:
            excel.new_document()
            target = Path(held) / 'k.xlsx'
            done = excel.run_vba(source, 'Make', (str(target),), timeout=120.0, line_numbers=False)
            assert done.ok and done.value == 'saved', done
            shutil.copyfile(target, BOOK)
    finally:
        os.rmdir(lock)
    print('saved', BOOK)


if __name__ == '__main__':
    make_book() if sys.argv[1] == 'book' else base.main(KINDS)
