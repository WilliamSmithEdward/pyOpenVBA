"""Range's most used methods held to the checklist's VERIFIED standard (docs/host_completeness.md): each with no argument
and with each kind of value for each argument, on a range of its own, late and early bound: the error, what it answers,
what the sheet then reads, and the same after a save and reopen.

    python scripts/measure_range_methods.py [group ...]

writes tests/fixtures/range_methods_*.json, one per group, which tests/test_excel_range_methods.py replays. Their "early"
holds the answers that differ with o typed. Copy, Cut and PasteSpecial use the clipboard, so this runs under excel-coordination/live.lock:
take it with mkdir first.
"""
from __future__ import annotations

import sys
from pathlib import Path

import measure_method_audit as audit

BLOCK = 'ws.Range("B2:B3")'
#: Values, a formula, text and a bold cell, with a cell out to one side.
_DATA = ('ws.Range("B2").Value = 1: ws.Range("B3").Value = 2: ws.Range("B4").Formula = "=B2+B3": ws.Range("C2").Value = "abc": '
         'ws.Range("C3").Value = "ABC": ws.Range("E6").Value = 9: ws.Range("B2").Font.Bold = True')
_CLEAR_CLIP = 'Application.CutCopyMode = False: ' + _DATA
_COPIED = _CLEAR_CLIP + ': ws.Range("B2:B4").Copy'
_MERGING = 'Application.DisplayAlerts = False: ' + _DATA
_MERGED = _MERGING + ': ws.Range("B2:C3").Merge'
_TABLE = ('ws.Range("B2:C2").Value = Array("h1", "h2"): ws.Range("B3:C3").Value = Array(3, "c"): ws.Range("B4:C4").Value = Array(1, "a"): '
          'ws.Range("B5:C5").Value = Array(2, "b")')
_FOUND = _DATA + ': ws.Range("B2:C4").Find "abc", , xlValues, xlPart, xlByRows, xlNext, False'
_OUTLINED = _DATA + ': ws.Rows("2:3").Group'
_HIDDEN = _DATA + ': ws.Rows(3).Hidden = True'
NUMBERS = ['1', '0', '-1', '99', '1.5', '"2"', '"x"', '""', 'Null', 'Empty', 'True']
#: Late bound, SpecialCells(Null) is automation error -2147417848, and Font and Interior then refuse every setting until
#: Excel starts again, so it is not measured.
NOT_NULL = [one for one in NUMBERS if one != 'Null']
BIG = ['1048577', '-1048577', '17000', '2147483648']
SHIFTS = ['xlShiftDown', 'xlShiftToRight', 'xlShiftUp', 'xlShiftToLeft', '2', '3', '4', *NUMBERS]
OBJECTS = ['ws.Range("G2")', '"G2"', 'Application', 'ws']
SWITCHES = ['True', 'False', '1', '0', '2', '"True"', '"x"', '""', 'Null', 'Empty']
TEXTS = ['"abc"', '"1"', '1', '"*"', '"?b?"', '"zzz"', '""', 'True', 'Null', 'Empty', 'ws.Range("C2")', 'ws.Range("C2:C3")', 'String(256, "a")']
LOOKS = ['xlValues', 'xlFormulas', 'xlComments', *NUMBERS]
PARTS = ['xlPart', 'xlWhole', *NUMBERS]
ORDERS = ['xlByRows', 'xlByColumns', *NUMBERS]
WAYS = ['xlNext', 'xlPrevious', *NUMBERS]
FIND = 'r = o.Find({}, {}, {}, {}, {}, {}, {}).Address'
READS = ['ws.Range("B2").Formula', 'ws.Range("B3").Formula', 'ws.Range("B4").Formula', 'ws.Range("C2").Formula', 'ws.Range("D2").Formula',
         'ws.Range("B2").Font.Bold', 'ws.UsedRange.Address']
PASTED = [*READS, 'ws.Range("G2").Formula', 'ws.Range("G3").Formula', 'ws.Range("G4").Formula', 'ws.Range("H2").Formula', 'ws.Range("G2").Font.Bold',
          'Application.CutCopyMode']
TABLE_READS = ['ws.Range("B3").Value', 'ws.Range("B4").Value', 'ws.Range("B5").Value', 'ws.Range("C3").Value', 'ws.Range("B2").Value',
               'ws.Rows(3).Hidden', 'ws.Rows(4).Hidden', 'ws.AutoFilterMode', 'ws.UsedRange.Address']
METHODS: list[tuple[str, str, str, str, str, list[str], list[tuple[str, list[str]]]]] = [
    ('Range', 'Offset', 'Range', 'ws.Range("C3")', '', [],
     [('r = o.Offset(ARG).Address', [*NUMBERS, *BIG]), ('r = o.Offset(0, ARG).Address', [*NUMBERS, *BIG]), ('r = o.Offset.Address', [])]),
    ('Range', 'Resize', 'Range', 'ws.Range("C3")', '', [],
     [('r = o.Resize(ARG).Address', [*NUMBERS, *BIG]), ('r = o.Resize(, ARG).Address', [*NUMBERS, *BIG]), ('r = o.Resize.Address', [])]),
    ('Range', 'End', 'Range', 'ws.Range("B2")', _DATA, [],
     [('r = o.End(ARG).Address', ['xlDown', 'xlUp', 'xlToLeft', 'xlToRight', '2', '3', '4', '5', *NUMBERS]), ('r = o.End.Address', [])]),
    ('Range', 'End', 'Range', 'ws.Range("B2")', _HIDDEN, [], [('r = o.End(ARG).Address', ['xlDown', 'xlToRight'])]),
    ('Range', 'End', 'Range', 'ws.Range("B2")', _MERGED, [], [('r = o.End(ARG).Address', ['xlDown', 'xlToRight'])]),
    ('Range', 'SpecialCells', 'Range', 'ws.Range("B2:C5")', _DATA, [],
     [('r = o.SpecialCells(ARG).Address', ['xlCellTypeConstants', 'xlCellTypeFormulas', 'xlCellTypeBlanks', 'xlCellTypeLastCell', 'xlCellTypeVisible',
                                           'xlCellTypeComments', 'xlCellTypeAllFormatConditions', 'xlCellTypeAllValidation', 'xlCellTypeSameFormatConditions',
                                           'xlCellTypeSameValidation', *NOT_NULL]),
      ('r = o.SpecialCells(xlCellTypeConstants, ARG).Address', ['xlNumbers', 'xlTextValues', 'xlLogical', 'xlErrors', *NOT_NULL]),
      ('r = o.SpecialCells.Address', [])]),
    ('Range', 'Clear', 'Range', BLOCK, _DATA, READS, [('o.Clear', []), ('r = o.Clear', []), ('o.Clear ARG', ['1'])]),
    ('Range', 'ClearContents', 'Range', BLOCK, _DATA, READS, [('o.ClearContents', []), ('r = o.ClearContents', [])]),
    ('Range', 'ClearOutline', 'Range', 'ws.Rows("2:3")', _OUTLINED, ['ws.Rows(2).OutlineLevel', 'ws.Rows(3).OutlineLevel'], [('o.ClearOutline', []), ('r = o.ClearOutline', [])]),
    ('Range', 'Insert', 'Range', BLOCK, _DATA, READS, [('o.Insert', []), ('r = o.Insert', []), ('o.Insert ARG', SHIFTS), ('o.Insert xlShiftDown, ARG', ['xlFormatFromLeftOrAbove', 'xlFormatFromRightOrBelow', *NUMBERS])]),
    ('Range', 'Insert', 'Range', 'ws.Rows(3)', _DATA, READS, [('o.Insert', []), ('o.Insert ARG', ['xlShiftDown', 'xlShiftToRight', '99', 'Null'])]),
    ('Range', 'Insert', 'Range', 'ws.Columns(3)', _DATA, READS, [('o.Insert', []), ('o.Insert ARG', ['xlShiftDown', 'xlShiftToRight', '99', 'Null'])]),
    ('Range', 'Delete', 'Range', BLOCK, _DATA, READS, [('o.Delete', []), ('r = o.Delete', []), ('o.Delete ARG', SHIFTS)]),
    ('Range', 'Delete', 'Range', 'ws.Rows(3)', _DATA, READS, [('o.Delete', []), ('o.Delete ARG', ['xlShiftUp', 'xlShiftToLeft', '99', 'Null'])]),
    ('Range', 'Delete', 'Range', 'ws.Columns(3)', _DATA, READS, [('o.Delete', []), ('o.Delete ARG', ['xlShiftUp', 'xlShiftToLeft', '99', 'Null'])]),
    ('Range', 'Merge', 'Range', 'ws.Range("B2:C3")', _MERGING, [*READS, 'ws.Range("B2").MergeArea.Address', 'ws.Range("C3").MergeCells'],
     [('o.Merge', []), ('r = o.Merge', []), ('o.Merge ARG', SWITCHES)]),
    ('Range', 'Merge', 'Range', 'ws.Range("C3")', _MERGING, ['ws.Range("C3").MergeCells'], [('o.Merge', [])]),
    ('Range', 'UnMerge', 'Range', 'ws.Range("B2:C3")', _MERGED, [*READS, 'ws.Range("B2").MergeCells'], [('o.UnMerge', []), ('r = o.UnMerge', []), ('o.UnMerge ARG', ['1'])]),
    ('Range', 'UnMerge', 'Range', 'ws.Range("C2")', _MERGED, ['ws.Range("B2").MergeCells'], [('o.UnMerge', [])]),
    ('Range', 'Copy', 'Range', 'ws.Range("B2:B4")', _CLEAR_CLIP, PASTED, [('o.Copy', []), ('r = o.Copy', []), ('o.Copy ARG', [*OBJECTS, *NUMBERS])]),
    ('Range', 'Cut', 'Range', 'ws.Range("B2:B4")', _CLEAR_CLIP, PASTED, [('o.Cut', []), ('r = o.Cut', []), ('o.Cut ARG', [*OBJECTS, *NUMBERS])]),
    ('Range', 'PasteSpecial', 'Range', 'ws.Range("G2")', _CLEAR_CLIP, PASTED, [('o.PasteSpecial', [])]),
    ('Range', 'PasteSpecial', 'Range', 'ws.Range("G2")', _COPIED, PASTED,
     [('o.PasteSpecial', []), ('r = o.PasteSpecial', []),
      ('o.PasteSpecial ARG', ['xlPasteAll', 'xlPasteValues', 'xlPasteFormulas', 'xlPasteFormats', 'xlPasteComments', 'xlPasteAllExceptBorders', 'xlPasteColumnWidths', '2', '3', '4', '5', '6', '7', '8', '9', '10', '11', '12', '13', *NUMBERS]),
      ('o.PasteSpecial xlPasteAll, ARG', ['xlPasteSpecialOperationNone', 'xlPasteSpecialOperationAdd', 'xlPasteSpecialOperationDivide', '2', '3', '4', '5', '6', *NUMBERS]),
      ('o.PasteSpecial xlPasteAll, xlNone, ARG', SWITCHES), ('o.PasteSpecial xlPasteAll, xlNone, False, ARG', SWITCHES)]),
    ('Range', 'Find', 'Range', 'ws.Range("B2:C4")', _DATA, [],
     [(FIND.format('ARG', '', 'xlValues', 'xlPart', 'xlByRows', 'xlNext', 'False'), TEXTS),
      (FIND.format('"abc"', 'ARG', 'xlValues', 'xlPart', 'xlByRows', 'xlNext', 'False'), ['ws.Range("B3")', 'ws.Range("Z9")', '"B3"', 'Null', 'Empty', '1', 'True']),
      (FIND.format('"abc"', '', 'ARG', 'xlPart', 'xlByRows', 'xlNext', 'False'), LOOKS),
      (FIND.format('"ab"', '', 'xlValues', 'ARG', 'xlByRows', 'xlNext', 'False'), PARTS),
      (FIND.format('"abc"', '', 'xlValues', 'xlPart', 'ARG', 'xlNext', 'False'), ORDERS),
      (FIND.format('"abc"', '', 'xlValues', 'xlPart', 'xlByRows', 'ARG', 'False'), WAYS),
      (FIND.format('"ABC"', '', 'xlValues', 'xlPart', 'xlByRows', 'xlNext', 'ARG'), SWITCHES),
      ('r = o.Find.Address', [])]),
    ('Range', 'FindNext', 'Range', 'ws.Range("B2:C4")', _FOUND, [],
     [('r = o.FindNext.Address', []), ('r = o.FindNext(ARG).Address', ['ws.Range("C2")', 'ws.Range("Z9")', '"C2"', 'Null', 'Empty', '1', 'True'])]),
    ('Range', 'FindPrevious', 'Range', 'ws.Range("B2:C4")', _FOUND, [],
     [('r = o.FindPrevious.Address', []), ('r = o.FindPrevious(ARG).Address', ['ws.Range("C2")', 'ws.Range("Z9")', '"C2"', 'Null', 'Empty', '1', 'True'])]),
    ('Range', 'Replace', 'Range', 'ws.Range("B2:C4")', _DATA, READS,
     [('r = o.Replace(ARG, "z", xlPart, xlByRows, False)', TEXTS),
      ('r = o.Replace("abc", ARG, xlPart, xlByRows, False)', TEXTS),
      ('r = o.Replace("ab", "z", ARG, xlByRows, False)', PARTS),
      ('r = o.Replace("abc", "z", xlPart, ARG, False)', ORDERS),
      ('r = o.Replace("ABC", "z", xlPart, xlByRows, ARG)', SWITCHES),
      ('r = o.Replace', [])]),
    ('Range', 'Sort', 'Range', 'ws.Range("B2:C5")', _TABLE, TABLE_READS,
     [('o.Sort', []), ('r = o.Sort(ws.Range("B2"), xlAscending, Header:=xlYes)', []),
      ('o.Sort ARG, Header:=xlYes', ['ws.Range("B2")', 'ws.Range("C2")', 'ws.Range("Z9")', '"B2"', '"h1"', '1', '0', 'Null', 'Empty', 'True']),
      ('o.Sort ws.Range("B2"), ARG, Header:=xlYes', ['xlAscending', 'xlDescending', *NUMBERS]),
      ('o.Sort ws.Range("B2"), xlAscending, Header:=ARG', ['xlYes', 'xlNo', 'xlGuess', *NUMBERS]),
      ('o.Sort ws.Range("B2"), xlAscending, Header:=xlYes, MatchCase:=ARG', SWITCHES),
      ('o.Sort ws.Range("B3"), xlAscending, Header:=xlYes, Orientation:=ARG', ['xlSortColumns', 'xlSortRows', *NUMBERS])]),
    ('Range', 'AutoFilter', 'Range', 'ws.Range("B2:C5")', _TABLE, TABLE_READS,
     [('o.AutoFilter', []), ('r = o.AutoFilter', []), ('o.AutoFilter ARG', NUMBERS),
      ('o.AutoFilter 1, ARG', ['3', '"3"', '">1"', '"<>3"', '"="', '""', 'True', 'Null', 'Empty', '"x"']),
      # xlFilterValues over one value disconnects Excel's objects, and the next statement raises an automation error.
      ('o.AutoFilter 1, 1, ARG', ['xlAnd', 'xlOr', 'xlTop10Items', *NUMBERS]),
      ('o.AutoFilter 1, ">1", xlAnd, ARG', ['"<3"', '3', '""', 'Null', 'Empty', 'True']),
      ('o.AutoFilter 1, 3, xlAnd, , ARG', SWITCHES)]),
]
#: What does not compile on a typed variable: a wrong number of arguments, a Sub's answer, and literal text for an enum.
WRONG_COUNT = (r'^o\.(Clear|UnMerge) 1$|^r = o\.(Merge|UnMerge)$|^r = o\.(End|SpecialCells|Find|Replace)(\.Address)?$'
               r'|o\.End\("|o\.SpecialCells\("|o\.PasteSpecial "|o\.PasteSpecial xlPasteAll, "'
               r'|o\.Sort ws\.Range\("B2"\), "|Header:="|Orientation:="|o\.AutoFilter 1, 1, "')


#: Each group is measured in its own run and fixture, as a module of every case is too large for Excel to load reliably.
GROUPS = {'range_methods_moves': ('Offset', 'Resize', 'End', 'SpecialCells', 'Clear', 'ClearContents', 'ClearOutline', 'Insert', 'Delete', 'Merge', 'UnMerge'),
          'range_methods_clipboard': ('Copy', 'Cut', 'PasteSpecial'),
          'range_methods_search': ('Find', 'FindNext', 'FindPrevious', 'Replace'),
          'range_methods_order': ('Sort', 'AutoFilter')}


def main() -> None:
    audit.WRONG_COUNT = WRONG_COUNT
    for name in sys.argv[1:] or GROUPS:
        audit.METHODS = [one for one in METHODS if one[1] in GROUPS[name]]
        audit.OUT = Path(__file__).resolve().parents[1] / f'tests/fixtures/{name}.json'
        audit.main()


if __name__ == '__main__':
    main()
