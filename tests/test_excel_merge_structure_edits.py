"""Merged areas following whole rows and columns, and cells, inserted and deleted, replayed from live Excel
(scripts/measure_merge_structure_edits.py, tests/fixtures/merge_structure_edits.json)."""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

RECORD: dict[str, Any] = json.loads(
    (Path(__file__).parent / 'fixtures/merge_structure_edits.json').read_text(encoding='utf-8'))


@pytest.fixture(scope='module')
def answers() -> list[str]:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(RECORD['source'], name='MergeEditsProbe')
    return str(app.run('MergeEditsProbe.Probe')).split('@@@')


@pytest.mark.parametrize('index', range(len(RECORD['cases'])))
def test_each_case_as_excel_answered(answers: list[str], index: int) -> None:
    assert answers[index] == RECORD['cases'][index][1]


def test_a_moved_merge_saves_where_it_went(tmp_path: Path) -> None:
    app = ExcelApplication()
    app.add_workbook()
    path = str(tmp_path / 'moved.xlsx')
    app.add_module('Public Function Trip(ByVal path As String) As String\n    Dim wb As Workbook\n'
                   '    Set wb = Workbooks.Add(xlWBATWorksheet)\n'
                   '    wb.Worksheets(1).Range("B2").Value = "m": wb.Worksheets(1).Range("B2:C3").Merge\n'
                   '    wb.Worksheets(1).Rows(3).Insert\n    wb.SaveAs path, xlOpenXMLWorkbook: wb.Close False\n'
                   '    Set wb = Workbooks.Open(path)\n'
                   '    Trip = wb.Worksheets(1).Range("B2").MergeArea.Address(0, 0) & "=" & wb.Worksheets(1).Range("B2")\n'
                   '    wb.Close False\nEnd Function\n', name='Trip')
    # Inserted through it, the merge grows a row (as the fixture's case).
    assert str(app.run('Trip.Trip', path)) == 'B2:C4=m'


def test_a_shift_that_unmerges_is_refused_while_excel_would_ask() -> None:
    # The probe turns DisplayAlerts off; with it on, Excel asks before it unmerges, which is not measured.
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('Public Sub Cut()\n    Range("B2:C3").Merge\n    Range("B1").Insert xlShiftDown\nEnd Sub\n',
                   name='Cut')
    with pytest.raises(VBAUnsupportedError, match='DisplayAlerts'):
        app.run('Cut.Cut')
