"""Set, read, save, reopen, read again, on Worksheet and the sheet's old collections, new and from a file (see set_save_reopen). What a sweep cannot say is whether a kept value matches Excel."""
from __future__ import annotations

import json
from pathlib import Path

import pytest
from set_save_reopen import Case, Target, cases, check

FIXTURES = Path(__file__).parent / 'fixtures'
BOOK = FIXTURES / 'old_collections.xlsx'
SHEET = {
    's.Visible': ['0', '2'], 's.StandardWidth': ['12'], 's.ScrollArea': ['"A1:C3"'], 's.EnableCalculation': ['False'], 's.DisplayPageBreaks': ['True'], 's.DisplayRightToLeft': ['True'],
    's.EnableAutoFilter': ['True'], 's.EnableOutlining': ['True'], 's.EnablePivotTable': ['True'], 's.EnableSelection': ['1'], 's.TransitionExpEval': ['True'], 's.TransitionFormEntry': ['True'],
    's.Name': ['"Renamed"'], 's.Tab.ColorIndex': ['6'], 's.Tab.Color': ['255'], 's.Tab.ThemeColor': ['4'], 's.Tab.TintAndShade': ['0.5'],
    's.Outline.SummaryRow': ['0'], 's.Outline.SummaryColumn': ['0'], 's.Outline.AutomaticStyles': ['True'],
    's.PageSetup.Orientation': ['2'], 's.PageSetup.Zoom': ['50'], 's.PageSetup.LeftMargin': ['20'], 's.PageSetup.RightMargin': ['21'], 's.PageSetup.TopMargin': ['22'], 's.PageSetup.BottomMargin': ['23'],
    's.PageSetup.HeaderMargin': ['24'], 's.PageSetup.FooterMargin': ['25'], 's.PageSetup.PaperSize': ['9'], 's.PageSetup.CenterHeader': ['"Head"'], 's.PageSetup.LeftFooter': ['"Foot"'],
    's.PageSetup.PrintGridlines': ['True'], 's.PageSetup.PrintHeadings': ['True'], 's.PageSetup.CenterHorizontally': ['True'], 's.PageSetup.CenterVertically': ['True'], 's.PageSetup.BlackAndWhite': ['True'],
    's.PageSetup.Draft': ['True'], 's.PageSetup.FirstPageNumber': ['3'], 's.PageSetup.Order': ['2'], 's.PageSetup.PrintArea': ['"$A$1:$B$2"'], 's.PageSetup.PrintTitleRows': ['"$1:$1"'],
    's.PageSetup.PrintTitleColumns': ['"$A:$A"'], 's.PageSetup.FitToPagesWide': ['2'], 's.PageSetup.FitToPagesTall': ['3'], 's.PageSetup.PrintComments': ['-4105'], 's.PageSetup.PrintErrors': ['1'],
    's.PageSetup.OddAndEvenPagesHeaderFooter': ['True'], 's.PageSetup.DifferentFirstPageHeaderFooter': ['True'], 's.PageSetup.ScaleWithDocHeaderFooter': ['False'], 's.PageSetup.AlignMarginsHeaderFooter': ['False'],
    's.Columns(2).ColumnWidth': ['15'], 's.Columns(2).Hidden': ['True'], 's.Rows(2).RowHeight': ['30'], 's.Rows(2).Hidden': ['True'], 's.Rows(2).OutlineLevel': ['2'], 's.Columns(2).OutlineLevel': ['2'],
}
ITEM = {
    's.Visible': ['False'], 's.Locked': ['False'], 's.PrintObject': ['False'], 's.Placement': ['3'], 's.Enabled': ['False'], 's.Caption': ['"New"'], 's.Text': ['"New"'], 's.LockedText': ['False'],
    's.Value': ['1'], 's.OnAction': ['"M"'], 's.Name': ['"Renamed"'], 's.Font.Bold': ['True'], 's.Font.Size': ['14'], 's.Font.ColorIndex': ['3'], 's.Interior.ColorIndex': ['6'], 's.Interior.Pattern': ['1'],
    's.Border.ColorIndex': ['3'], 's.Border.Weight': ['2'], 's.Border.LineStyle': ['1'], 's.LinkedCell': ['"$D$1"'], 's.ListFillRange': ['"$A$20:$A$22"'], 's.DropDownLines': ['4'], 's.MultiSelect': ['2'],
    's.Display3DShading': ['True'], 's.HorizontalAlignment': ['-4108'], 's.VerticalAlignment': ['-4108'], 's.Orientation': ['-4171'], 's.AutoSize': ['True'], 's.RoundedCorners': ['True'], 's.Shadow': ['True'],
    's.ArrowHeadStyle': ['2'], 's.AddIndent': ['True'], 's.DefaultButton': ['True'], 's.CancelButton': ['True'], 's.Accelerator': ['"N"'], 's.Left': ['30'], 's.Top': ['31'], 's.Width': ['60'], 's.Height': ['45'],
}
ADDS = {'TextBoxes': '10, 10, 100, 40', 'Rectangles': '10, 10, 100, 40', 'Ovals': '10, 10, 100, 40', 'Arcs': '10, 10, 100, 40', 'Lines': '10, 10, 100, 40', 'Buttons': '10, 10, 100, 40', 'CheckBoxes': '10, 10, 100, 20',
        'OptionButtons': '10, 10, 100, 20', 'DropDowns': '10, 10, 100, 20', 'ListBoxes': '10, 10, 100, 60'}
SHEETS = [Target('sheet new', 'Set s = ws', 'Set s = ws'), Target('sheet file', 'Set s = ws', 'Set s = ws', BOOK)]
ITEMS = ([Target(f'{name} new', f'ws.Range("A20").Value = "a": ws.Range("A21").Value = "b": Set s = ws.{name}.Add({box})', f'Set s = ws.{name}(1)') for name, box in ADDS.items()]
         + [Target(f'{name} file', f'Set s = ws.{name}(1)', f'Set s = ws.{name}(1)', BOOK) for name in (*ADDS, 'Pictures')])
CASES: list[Case] = cases(SHEETS, SHEET) + cases(ITEMS, ITEM)
#: What the file does not keep yet.
ACCEPTED = {f'{name} new: s.Top = 31': 'Set on a new control it reads as set, and from the file at the pixel its box snapped to, 30.75; the set does not snap yet'
            for name in ('Buttons', 'CheckBoxes', 'OptionButtons', 'DropDowns', 'ListBoxes')}
_MEASURED = json.loads((FIXTURES / 'sheet_session_settings.json').read_text(encoding='utf-8'))['cases']
#: What Excel's own save loses too, hidden and shown, so a model keeping it would be wrong.
EXCEL_LOSES = {f'{target.name}: s.{one["member"]} = {one["value"]}': 'tests/fixtures/sheet_session_settings.json' for one in _MEASURED if not one['shown'] for target in SHEETS}


@pytest.mark.parametrize('case', CASES, ids=str)
def test_each_setting_comes_back(case: Case, tmp_path: Path) -> None:
    check(case, tmp_path, ACCEPTED, EXCEL_LOSES)


def test_excel_lost_each_of_them_either_way() -> None:
    assert len(_MEASURED) == 12
    assert all(one['before'] != one['after'] for one in _MEASURED)
