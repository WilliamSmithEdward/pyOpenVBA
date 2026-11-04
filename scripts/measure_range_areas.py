"""Range's methods on ranges of several areas, and the other cases they report as gaps: what the error is, what it answers,
what the sheet then reads, and the same after a save and reopen, late and early bound.

    python scripts/measure_range_areas.py

writes tests/fixtures/range_areas.json, which tests/test_excel_range_areas.py replays. Its "early" holds the answers
that differ with o typed.
"""
from __future__ import annotations

from pathlib import Path

import measure_method_audit as audit

_GRID = ('ws.Range("B2").Value = "abc": ws.Range("B3").Value = 2: ws.Range("B4").Value = 3: ws.Range("C2").Value = 4: ws.Range("C3").Value = 5: '
         'ws.Range("C4").Value = 6: ws.Range("D2").Value = 7: ws.Range("D3").Value = "abc": ws.Range("D4").Value = 9: ws.Range("D5").Value = 10: '
         'ws.Range("E2").Value = 11: ws.Range("E3").Value = 12: ws.Range("E4").Value = "abc": ws.Range("E5").Formula = "=B3+C3"')
_MERGING = 'Application.DisplayAlerts = False: ' + _GRID
_MERGED = _MERGING + ': ws.Range("B2:C3").Merge: ws.Range("E4:E5").Merge'
_NOTES = _GRID + ': ws.Range("B3").AddComment "abc one": ws.Range("D3").AddComment "two"'
_FORMATS = (_MERGING + ': ws.Range("B2").Font.Bold = True: ws.Range("B2").Interior.Color = 255: ws.Range("C2").NumberFormat = "0.00": '
            'ws.Range("B3").Font.Italic = True: ws.Range("C3").Borders(xlEdgeBottom).LineStyle = xlContinuous')
FORMAT_READS = [f'ws.Range("{cell}").{member}' for cell in ('B2', 'C2', 'B3', 'C3')
                for member in ('Font.Bold', 'Font.Italic', 'Interior.Color', 'NumberFormat', 'Borders(xlEdgeBottom).LineStyle', 'MergeCells')]
_BORDERED = (_MERGING + ': ws.Range("B2").Borders(xlEdgeLeft).LineStyle = xlContinuous: ws.Range("B2").Borders(xlEdgeTop).LineStyle = xlDash: '
             'ws.Range("B2").Borders(xlEdgeRight).LineStyle = xlDot: ws.Range("B2").Borders(xlEdgeBottom).LineStyle = xlDouble: '
             'ws.Range("C3").Borders(xlEdgeRight).LineStyle = xlContinuous')
_EDGES = (_MERGING + ': ws.Range("B2").Borders(xlEdgeLeft).LineStyle = xlDash: ws.Range("B2").Borders(xlEdgeTop).LineStyle = xlContinuous: '
          'ws.Range("C2").Borders(xlEdgeRight).LineStyle = xlDouble: ws.Range("C2").Borders(xlEdgeBottom).LineStyle = xlContinuous: '
          'ws.Range("B3").Borders(xlEdgeLeft).LineStyle = xlContinuous: ws.Range("B3").Borders(xlEdgeBottom).LineStyle = xlDash')
BORDER_READS = [f'ws.Range("{cell}").Borders({edge}).LineStyle' for cell in ('B2', 'C2', 'B3', 'C3')
                for edge in ('xlEdgeLeft', 'xlEdgeTop', 'xlEdgeRight', 'xlEdgeBottom')]
GRID = [f'ws.Range("{cell}").Formula' for cell in ('B2', 'B3', 'B4', 'C2', 'C3', 'C4', 'D2', 'D3', 'D4', 'D5', 'E2', 'E3', 'E4', 'E5', 'F2', 'F3', 'B5', 'C5')]
READS = [*GRID, 'ws.UsedRange.Address']
MERGE_READS = [*READS, 'ws.Range("B2").MergeArea.Address', 'ws.Range("D3").MergeArea.Address', 'ws.Range("E4").MergeArea.Address']
TWO = 'ws.Range("B2:B3,D3:D4")'
SIDE = 'ws.Range("B2:B3,E2:E3")'
ROW = 'ws.Range("B2:C2,B4:C4")'
OVER = 'ws.Range("B2:C3,C3:D4")'
ROWS = 'ws.Range("2:2,4:4")'
COLUMNS = 'ws.Range("B:B,D:D")'
MIXED = 'ws.Range("2:2,D:D")'
FIRST = 'r = o.Find("abc", , xlValues, xlPart, xlByRows, xlNext, False).Address'
BACK = 'r = o.Find("abc", , xlValues, xlPart, xlByRows, xlPrevious, False).Address'
METHODS: list[tuple[str, str, str, str, str, list[str], list[tuple[str, list[str]]]]] = [
    *[('Range', 'Merge', 'Range', target, _MERGING, MERGE_READS, [('o.Merge', []), ('o.Merge True', [])])
      for target in ('ws.Range("B2:C3,D3:E4")', TWO, OVER)],
    ('Range', 'Merge', 'Range', 'ws.Range("B2:C3")', _MERGED, MERGE_READS, [('o.Merge True', [])]),
    ('Range', 'Merge', 'Range', 'ws.Range("B2:E5")', _MERGED, MERGE_READS, [('o.Merge', []), ('o.Merge True', [])]),
    # What a merge does to the cells' formats: the first cell's goes over them all, and unmerging leaves it there.
    ('Range', 'Merge', 'Range', 'ws.Range("B2:C3")', _FORMATS, FORMAT_READS, [('o.Merge', []), ('o.Merge True', [])]),
    ('Range', 'UnMerge', 'Range', 'ws.Range("B2:C3")', _FORMATS + ': ws.Range("B2:C3").Merge', FORMAT_READS, [('o.UnMerge', [])]),
    ('Range', 'Merge', 'Range', 'ws.Range("B2:C3")', _BORDERED, BORDER_READS, [('o.Merge', []), ('o.Merge True', [])]),
    ('Range', 'UnMerge', 'Range', 'ws.Range("B2:C3")', _BORDERED + ': ws.Range("B2:C3").Merge', BORDER_READS, [('o.UnMerge', [])]),
    # Which edges stay is told apart from which line it was: dashes outside, continuous lines on top and below.
    ('Range', 'Merge', 'Range', 'ws.Range("B2:C3")', _EDGES, BORDER_READS, [('o.Merge True', []), ('o.Merge', [])]),
    ('Range', 'Merge', 'Range', 'ws.Range("B2:C2")', _EDGES, BORDER_READS, [('o.Merge', []), ('o.Merge True', [])]),
    ('Range', 'Merge', 'Range', 'ws.Range("B2:B3")', _EDGES, BORDER_READS, [('o.Merge', []), ('o.Merge True', [])]),
    *[('Range', 'UnMerge', 'Range', target, _MERGED, MERGE_READS, [('o.UnMerge', [])])
      for target in ('ws.Range("B2:C3,E4:E5")', 'ws.Range("B2,E4")', 'ws.Range("B2,D2")', 'ws.Range("B2:E5")')],
    *[('Range', 'Insert', 'Range', target, _GRID, READS, [('o.Insert', []), ('o.Insert xlShiftDown', []), ('o.Insert xlShiftToRight', [])])
      for target in (TWO, SIDE, ROW, OVER)],
    *[('Range', 'Insert', 'Range', target, _GRID, READS, [('o.Insert', [])]) for target in (ROWS, COLUMNS, MIXED)],
    *[('Range', 'Delete', 'Range', target, _GRID, READS, [('o.Delete', []), ('o.Delete xlShiftUp', []), ('o.Delete xlShiftToLeft', [])])
      for target in (TWO, SIDE, ROW, OVER)],
    *[('Range', 'Delete', 'Range', target, _GRID, READS, [('o.Delete', [])]) for target in (ROWS, COLUMNS, MIXED)],
    *[('Range', 'Find', 'Range', target, _GRID, [], [(FIRST, []), (BACK, [])]) for target in (TWO, SIDE, 'ws.Range("B4,E4,D3")', 'ws.Range("C2:C3,B2")')],
    # Area order and sheet order answer differently here. _ALL has "abc" in each of B2, D3 and E4.
    *[('Range', 'Find', 'Range', target, _GRID, [],
       [(FIRST, []), (BACK, []), ('r = o.Find("abc", , xlValues, xlPart, xlByColumns, xlNext, False).Address', []),
        ('r = o.Find("abc", , xlValues, xlPart, xlByColumns, xlPrevious, False).Address', []),
        ('r = o.Find("abc", ws.Range("D3"), xlValues, xlPart, xlByRows, xlNext, False).Address', [])])
      for target in ('ws.Range("E4,D3,B2")', 'ws.Range("D3,E4,B2")', 'ws.Range("D3,B2,E4")', 'ws.Range("E3:E4,B2:B3,D2:D3")')],
    ('Range', 'FindNext', 'Range', TWO, _GRID + f': {TWO}.Find "abc", , xlValues, xlPart, xlByRows, xlNext, False', [],
     [('r = o.FindNext.Address', []), ('r = o.FindNext(ws.Range("B2")).Address', []), ('r = o.FindNext(ws.Range("D3")).Address', []),
      ('r = o.FindNext(ws.Range("D4")).Address', []), ('r = o.FindNext(ws.Range("C2")).Address', [])]),
    ('Range', 'FindPrevious', 'Range', TWO, _GRID + f': {TWO}.Find "abc", , xlValues, xlPart, xlByRows, xlNext, False', [],
     [('r = o.FindPrevious.Address', []), ('r = o.FindPrevious(ws.Range("B2")).Address', []), ('r = o.FindPrevious(ws.Range("D3")).Address', []),
      ('r = o.FindPrevious(ws.Range("D4")).Address', []), ('r = o.FindPrevious(ws.Range("C2")).Address', [])]),
    ('Range', 'Find', 'Range', 'ws.Range("B2:D4")', _NOTES, [],
     [('r = o.Find("abc", , xlComments, xlPart, xlByRows, xlNext, False).Address', []), ('r = o.Find("two", , xlComments, xlWhole, xlByRows, xlNext, False).Address', []),
      ('r = o.Find("TWO", , xlComments, xlPart, xlByRows, xlNext, True).Address', []), ('r = o.Find("abc", , xlComments, xlPart, xlByRows, xlPrevious, False).Address', []),
      ('r = o.Find("zzz", , xlComments, xlPart, xlByRows, xlNext, False).Address', [])]),
    ('Range', 'Find', 'Range', 'ws.Range("B2:D4")', _GRID, [],
     [('r = o.Find("abc", , xlValues, xlPart, xlByRows, xlNext, False, True).Address', []), ('r = o.Find("abc", , xlValues, xlPart, xlByRows, xlNext, False, False).Address', [])]),
    ('Range', 'Replace', 'Range', TWO, _GRID, READS, [('r = o.Replace("abc", "z", xlPart, xlByRows, False)', [])]),
    ('Range', 'Replace', 'Range', 'ws.Range("B2:D4")', _GRID, READS, [('r = o.Replace("abc", "z", xlPart, xlByRows, False, True)', []), ('r = o.Replace("abc", "z", xlPart, xlByRows, False, False)', [])]),
]


def main() -> None:
    audit.METHODS = METHODS
    audit.WRONG_COUNT = r'^$'
    audit.OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/range_areas.json'
    audit.main()


if __name__ == '__main__':
    main()
