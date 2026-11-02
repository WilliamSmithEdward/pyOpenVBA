"""Window.SheetViews and ActiveSheetView, replayed from live Excel (scripts/measure_sheet_views.py,
tests/fixtures/sheet_views.json)."""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/sheet_views.json').read_text(encoding='utf-8'))


@pytest.fixture(scope='module')
def answers() -> list[str]:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(RECORD['source'], name='ViewsProbe')
    return str(app.run('ViewsProbe.Probe')).split('@@@')


@pytest.mark.parametrize('index', range(len(RECORD['cases'])), ids=[case for case, _ in RECORD['cases']])
def test_each_case_as_excel_answered(answers: list[str], index: int) -> None:
    assert answers[index] == RECORD['cases'][index][1]


def test_a_view_set_while_its_sheet_is_inactive_is_saved_with_that_sheet(tmp_path: Path) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('Public Sub Go(ByVal path As String)\n    ActiveWorkbook.Worksheets.Add After:=ActiveSheet\n'
                   '    ActiveWindow.SheetViews(1).DisplayGridlines = False\n'
                   '    ActiveWorkbook.SaveAs path, xlOpenXMLWorkbook\nEnd Sub\n', name='Views')
    app.run('Views.Go', str(tmp_path / 'views.xlsx'))
    reopened = ExcelApplication().open_workbook(str(tmp_path / 'views.xlsx'))
    assert [sheet.view.gridlines for sheet in reopened.sheets_] == [False, True]
