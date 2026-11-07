"""Set, read, save, reopen, read again, on tables and pivot tables, new and from a file Excel saved (see set_save_reopen). What a sweep cannot
say is whether a kept value matches Excel. ListRow has no member that can be set, so it has no cases."""
from __future__ import annotations

import json
import re
from pathlib import Path

import pytest
from set_save_reopen import Case, Target, cases, check

FIXTURES = Path(__file__).parent / 'fixtures'
BOOK = FIXTURES / 'dxf_table' / 'made.xlsx'
DATA = 'ws.Range("A1:B1").Value = Array("Name", "Amount"): ws.Range("A2:B4").Value = ws.Evaluate("{""a"",1;""b"",2;""a"",3}")'
NEW_TABLE = f'{DATA}: Set s = ws.ListObjects.Add(1, ws.Range("A1:B4"), , 1): s.Name = "T"'
NEW_PIVOT = (f'{DATA}: Set s = ActiveWorkbook.PivotCaches.Create(1, ws.Name & "!R1C1:R4C2").CreatePivotTable(ws.Range("D3"), "P"): '
             's.PivotFields("Name").Orientation = 1: s.AddDataField s.PivotFields("Amount")')
FILE_PIVOT = 'Set s = ActiveWorkbook.Worksheets("pivot").PivotTables(1)'

TABLES = [Target('table new', NEW_TABLE, 'Set s = ws.ListObjects(1)'), Target('table file', 'Set s = ws.ListObjects(1)', 'Set s = ws.ListObjects(1)', BOOK)]
TABLE = {
    's.Name': ['"Renamed"'], 's.DisplayName': ['"Shown"'], 's.AlternativeText': ['"Alt"'], 's.Summary': ['"Sum"'], 's.Comment': ['"Note"'],
    's.ShowAutoFilter': ['False'], 's.ShowAutoFilterDropDown': ['False'], 's.ShowHeaders': ['False'], 's.ShowTotals': ['True'],
    's.ShowTableStyleColumnStripes': ['True'], 's.ShowTableStyleFirstColumn': ['True'], 's.ShowTableStyleLastColumn': ['True'], 's.ShowTableStyleRowStripes': ['False'],
    's.TableStyle': ['"TableStyleMedium9"', '""'],
}
COLUMNS = [Target('column new', f'{NEW_TABLE}: Set s = s.ListColumns(2)', 'Set s = ws.ListObjects(1).ListColumns(2)'),
           Target('column file', 'Set s = ws.ListObjects(1).ListColumns(2)', 'Set s = ws.ListObjects(1).ListColumns(2)', BOOK)]
COLUMN = {'s.Name': ['"Renamed"'], 's.TotalsCalculation': ['1', '2', '5']}

PIVOTS = [Target('pivot new', NEW_PIVOT, 'Set s = ws.PivotTables(1)'), Target('pivot file', FILE_PIVOT, FILE_PIVOT, BOOK)]
PIVOT = {
    's.Name': ['"Renamed"'], 's.ColumnGrand': ['False'], 's.RowGrand': ['False'], 's.GrandTotalName': ['"All"'], 's.ErrorString': ['"err"'], 's.DisplayErrorString': ['True'],
    's.NullString': ['"none"'], 's.DisplayNullString': ['False'], 's.MergeLabels': ['True'], 's.PreserveFormatting': ['False'], 's.SaveData': ['False'], 's.EnableDrilldown': ['False'],
    's.EnableFieldList': ['False'], 's.EnableWizard': ['False'], 's.EnableFieldDialog': ['False'], 's.EnableDataValueEditing': ['True'], 's.DisplayFieldCaptions': ['False'],
    's.DisplayContextTooltips': ['False'], 's.ShowDrillIndicators': ['False'], 's.PrintDrillIndicators': ['True'], 's.PrintTitles': ['True'], 's.RepeatItemsOnEachPrintedPage': ['False'],
    's.InGridDropZones': ['True'], 's.FieldListSortAscending': ['True'], 's.SortUsingCustomLists': ['False'], 's.SubtotalHiddenPageItems': ['True'], 's.ShowPageMultipleItemLabel': ['False'],
    's.PageFieldOrder': ['1'], 's.PageFieldWrapCount': ['2'], 's.CompactRowIndent': ['3'], 's.CompactLayoutRowHeader': ['"Rows"'], 's.CompactLayoutColumnHeader': ['"Cols"'],
    's.DisplayEmptyRow': ['True'], 's.DisplayEmptyColumn': ['True'], 's.DisplayImmediateItems': ['False'], 's.HasAutoFormat': ['False'], 's.ManualUpdate': ['True'],
    's.ShowTableStyleColumnHeaders': ['False'], 's.ShowTableStyleRowHeaders': ['False'], 's.ShowTableStyleColumnStripes': ['True'], 's.ShowTableStyleRowStripes': ['True'],
    's.ShowTableStyleLastColumn': ['True'], 's.ShowValuesRow': ['True'], 's.TableStyle2': ['"PivotStyleMedium9"'], 's.AlternativeText': ['"Alt"'], 's.Summary': ['"Sum"'],
    's.Tag': ['"tag"'], 's.VacatedStyle': ['"Normal"'], 's.AllowMultipleFilters': ['False'], 's.CalculatedMembersInFilters': ['True'], 's.VisualTotals': ['False'],
    's.TotalsAnnotation': ['True'], 's.ViewCalculatedMembers': ['False'], 's.LayoutRowDefault': ['1'], 's.SmallGrid': ['False'],
}
FIELDS = [Target('field new', f'{NEW_PIVOT}: Set s = s.PivotFields("Name")', 'Set s = ws.PivotTables(1).PivotFields(1)'),
          Target('field file', f'{FILE_PIVOT}: Set s = s.PivotFields("Name")', f'{FILE_PIVOT}: Set s = s.PivotFields(1)', BOOK)]
FIELD = {
    's.Caption': ['"Who"'], 's.NumberFormat': ['"0.00"'], 's.LayoutBlankLine': ['True'], 's.LayoutCompactRow': ['False'], 's.LayoutForm': ['0'], 's.LayoutPageBreak': ['True'],
    's.LayoutSubtotalLocation': ['2'], 's.RepeatLabels': ['True'], 's.ShowAllItems': ['True'], 's.SubtotalName': ['"Sub"'], 's.Subtotals(1)': ['False'], 's.Subtotals(2)': ['True'],
    's.DragToColumn': ['False'], 's.DragToData': ['False'], 's.DragToHide': ['False'], 's.DragToPage': ['False'], 's.DragToRow': ['False'], 's.DatabaseSort': ['True'],
    's.EnableItemSelection': ['False'], 's.IncludeNewItemsInFilter': ['True'], 's.Orientation': ['2', '3', '0'], 's.Position': ['1'],
}
#: A renamed item is found again by its place, a moved one by its name.
ITEMS = [Target('item new', f'{NEW_PIVOT}: Set s = s.PivotFields("Name").PivotItems("b")', 'Set s = ws.PivotTables(1).PivotFields(1).PivotItems(2)', kind='renamed'),
         Target('item new', f'{NEW_PIVOT}: Set s = s.PivotFields("Name").PivotItems("b")', 'Set s = ws.PivotTables(1).PivotFields(1).PivotItems("b")', kind='moved'),
         Target('item file', f'{FILE_PIVOT}: Set s = s.PivotFields("Name").PivotItems(1)', f'{FILE_PIVOT}: Set s = s.PivotFields("Name").PivotItems(1)', BOOK, kind='renamed'),
         Target('item file', f'{FILE_PIVOT}: Set s = s.PivotFields("Name").PivotItems(1)', f'{FILE_PIVOT}: Set s = s.PivotFields("Name").PivotItems(1)', BOOK, kind='moved')]
ITEM = {'s.Caption': ['"Bee"'], 's.Visible': ['False'], 's.Position': ['1'], 's.ShowDetail': ['False']}
ITEM_KINDS = {'s.Caption': ('renamed',), 's.Visible': ('moved',), 's.Position': ('moved',), 's.ShowDetail': ('moved',)}

CASES: list[Case] = cases(TABLES, TABLE) + cases(COLUMNS, COLUMN) + cases(PIVOTS, PIVOT) + cases(FIELDS, FIELD) + cases(ITEMS, ITEM, ITEM_KINDS)
#: What the file does not keep yet.
ACCEPTED: dict[str, str] = {}
_OPTIONS = json.loads((FIXTURES / 'pivot_options' / 'pivot_options.json').read_text(encoding='utf-8'))['layouts']
#: What Excel's own save loses too: ManualUpdate set on reads True, and False once Excel reopens its own save.
EXCEL_LOSES = {f'{target.name}: s.ManualUpdate = True': 'tests/fixtures/pivot_options/pivot_options.json, the manual_ layouts' for target in PIVOTS}


@pytest.mark.parametrize('case', CASES, ids=str)
def test_each_setting_comes_back(case: Case, tmp_path: Path) -> None:
    check(case, tmp_path, ACCEPTED, EXCEL_LOSES)


def test_excel_loses_manual_update_across_its_own_save() -> None:
    held = [re.findall(r'ManualUpdate=(\w+)', json.dumps(layout)) for layout in _OPTIONS if layout['name'].startswith('manual_') and layout['name'] != 'manual_done']
    assert len(held) == 8
    assert all(reads == ['True', 'False'] for reads in held)
