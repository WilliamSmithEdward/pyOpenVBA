"""Set, read, save, reopen, read again, on the held families no other sweep reaches: hyperlinks, notes, defined names and conditional
formats, new and from a file Excel saved (see set_save_reopen). What a sweep cannot say is whether a kept value matches Excel. The file's
third and fourth links, a mail link and a shape's, are left alone, as is EmailSubject, and a format's fonts are Font's own sweep."""
from __future__ import annotations

import json
from pathlib import Path

import pytest
from set_save_reopen import Case, Target, cases, check

FIXTURES = Path(__file__).parent / 'fixtures'
BOOK = FIXTURES / 'object_roundtrip.xlsx'

LINKS = [Target('link new', 'Set s = ws.Hyperlinks.Add(ws.Range("C3"), "http://example.com/x", "", "", "Go")', 'Set s = ws.Range("C3").Hyperlinks(1)'),
         Target('link inner new', 'Set s = ws.Hyperlinks.Add(ws.Range("C4"), "", "Sheet1!A1", "", "In")', 'Set s = ws.Range("C4").Hyperlinks(1)'),
         Target('link file', 'Set s = ws.Range("B10").Hyperlinks(1)', 'Set s = ws.Range("B10").Hyperlinks(1)', BOOK),
         Target('link inner file', 'Set s = ws.Range("B11").Hyperlinks(1)', 'Set s = ws.Range("B11").Hyperlinks(1)', BOOK)]
LINK = {'s.Address': ['"http://example.org/y"'], 's.SubAddress': ['"Sheet1!B2"', '""'], 's.ScreenTip': ['"Tip"'], 's.TextToDisplay': ['"Other"']}

NOTES = [Target('note new', 'Set s = ws.Range("C3").AddComment("Hello")', 'Set s = ws.Range("C3").Comment'),
         Target('note file', 'Set s = ws.Range("A10").Comment', 'Set s = ws.Range("A10").Comment', BOOK),
         Target('note hidden file', 'Set s = ws.Range("A11").Comment', 'Set s = ws.Range("A11").Comment', BOOK)]
NOTE = {'s.Visible': ['True', 'False'], 's.Author': ['"Someone"'], 's.Shape.Width': ['150'], 's.Shape.Height': ['80'], 's.Shape.Left': ['200'], 's.Shape.Top': ['40'],
        's.Shape.Fill.ForeColor.RGB': ['65535'], 's.Shape.TextFrame.AutoSize': ['True'], 's.Shape.Placement': ['1']}

NAMES = [Target('name new', 'Set s = ActiveWorkbook.Names.Add("Held", "=Sheet1!$A$1")', 'Set s = ActiveWorkbook.Names("Held")'),
         Target('name file', 'Set s = ActiveWorkbook.Names("Commented")', 'Set s = ActiveWorkbook.Names("Commented")', BOOK),
         Target('name constant file', 'Set s = ActiveWorkbook.Names("Constant")', 'Set s = ActiveWorkbook.Names("Constant")', BOOK)]
NAME = {'s.RefersTo': ['"=$B$2:$C$3"', '"=7"', '"=SUM(1,2)"'], 's.RefersToR1C1': ['"=R5C5"'], 's.Comment': ['"Why"', '""'], 's.Visible': ['False'], 's.Value': ['"=9"'],
        's.MacroType': ['1'], 's.ShortcutKey': ['"q"'], 's.Category': ['"Mine"'], 's.WorkbookParameter': ['True']}

FORMATS = [Target('format new', 'ws.Range("D1:D3").Value = 5: Set s = ws.Range("D1:D3").FormatConditions.Add(1, 5, "1", "9")', 'Set s = ws.Range("D1:D3").FormatConditions(1)'),
           Target('format text new', 'Set s = ws.Range("E1:E3").FormatConditions.Add(Type:=9, String:="abc", TextOperator:=0)', 'Set s = ws.Range("E1:E3").FormatConditions(1)', kind='text'),
           Target('format file', 'Set s = ws.Range("D10:D12").FormatConditions(1)', 'Set s = ws.Range("D10:D12").FormatConditions(1)', BOOK)]
FORMAT = {'s.StopIfTrue': ['True'], 's.Priority': ['1'], 's.NumberFormat': ['"0.000"'], 's.Interior.Color': ['255'], 's.Interior.Pattern': ['9'], 's.Borders(1).LineStyle': ['1'],
          's.Text': ['"xyz"'], 's.TextOperator': ['1', '3']}
FORMAT_KINDS = {'s.Text': ('text',), 's.TextOperator': ('text',)}

#: Held in a variable across an edit that moves it, then set through the variable: the set must land where the object now is. One held
#: across an edit that takes it away is set too: that set raises 424 and claims nothing, so there is nothing a save could lose.
TABLE_DATA = 'ws.Range("B2:C2").Value = Array("x", "y"): ws.Range("B3:C4").Value = 1'
HELD = [Target('range held moved', 'Set s = ws.Range("C3"): ws.Rows(1).Insert: ws.Columns(1).Insert', 'Set s = ws.Range("D4")', kind='range'),
        Target('range held moved file', 'Set s = ws.Range("C3"): ws.Rows(1).Insert', 'Set s = ws.Range("C4")', BOOK, kind='range'),
        Target('range held gone', 'Set s = ws.Range("C3"): ws.Rows(3).Delete', 'Set s = ws.Range("C3")', kind='range'),
        Target('table held moved', f'{TABLE_DATA}: Set s = ws.ListObjects.Add(1, ws.Range("B2:C4"), , 1): ws.Rows(1).Insert: ws.Columns(1).Insert', 'Set s = ws.ListObjects(1)', kind='table'),
        Target('table held gone', f'{TABLE_DATA}: Set s = ws.ListObjects.Add(1, ws.Range("B2:C4"), , 1): ws.Rows("2:4").Delete', 'Set s = ws.ListObjects(1)', kind='table'),
        Target('shape held moved', 'Set s = ws.Shapes.AddShape(1, 100, 100, 50, 50): ws.Columns(1).Insert: ws.Rows(1).Insert', 'Set s = ws.Shapes(1)', kind='shape'),
        Target('note held moved', 'Set s = ws.Range("C3").AddComment("Held"): ws.Rows(1).Insert', 'Set s = ws.Range("C4").Comment', kind='note'),
        Target('link held moved', 'Set s = ws.Hyperlinks.Add(ws.Range("C3"), "http://example.com/x"): ws.Columns(1).Insert', 'Set s = ws.Range("D3").Hyperlinks(1)', kind='link')]
HELD_SETTINGS = {'s.Value': ['7', '"text"'], 's.NumberFormat': ['"0.0%"'], 's.Interior.ColorIndex': ['6'], 's.HorizontalAlignment': ['-4108'],
                 's.Name': ['"Kept"'], 's.ShowTotals': ['True'], 's.TableStyle': ['"TableStyleLight9"'], 's.Rotation': ['30'], 's.AlternativeText': ['"Alt"'],
                 's.Visible': ['True'], 's.ScreenTip': ['"Tip"'], 's.TextToDisplay': ['"Other"']}
HELD_KINDS = {'s.Value': ('range',), 's.NumberFormat': ('range',), 's.Interior.ColorIndex': ('range',), 's.HorizontalAlignment': ('range',),
              's.Name': ('table', 'shape'), 's.ShowTotals': ('table',), 's.TableStyle': ('table',), 's.Rotation': ('shape',), 's.AlternativeText': ('table', 'shape'),
              's.Visible': ('note',), 's.ScreenTip': ('link',), 's.TextToDisplay': ('link',)}

CASES: list[Case] = (cases(LINKS, LINK) + cases(NOTES, NOTE) + cases(NAMES, NAME) + cases(FORMATS, FORMAT, FORMAT_KINDS)
                     + cases(HELD, HELD_SETTINGS, HELD_KINDS))
#: What the file does not keep yet.
ACCEPTED: dict[str, str] = {}
_MACRO_NAMES = json.loads((FIXTURES / 'name_macro.json').read_text(encoding='utf-8'))['names']
#: What Excel's own save loses too: a workbook saved without macros leaves an Excel 4 function name out.
EXCEL_LOSES = {f'{target.name}: s.MacroType = 1': 'tests/fixtures/name_macro.json, function.xlsx' for target in NAMES}


@pytest.mark.parametrize('case', CASES, ids=str)
def test_each_setting_comes_back(case: Case, tmp_path: Path) -> None:
    check(case, tmp_path, ACCEPTED, EXCEL_LOSES)


def test_an_object_gone_after_the_reopen_fails_unless_listed(tmp_path: Path) -> None:
    gone = next(case for case in CASES if str(case) == 'name new: s.MacroType = 1')
    with pytest.raises(AssertionError, match='is lost across a save: 1 before the save; after it the object is not found again'):
        check(gone, tmp_path, {})
    check(gone, tmp_path, {str(gone): 'accepted'})


def test_excel_leaves_a_function_name_out_of_a_book_without_macros() -> None:
    assert _MACRO_NAMES['function.xlsx'] == ''
    assert 'function="1"' in _MACRO_NAMES['function.xlsm']
