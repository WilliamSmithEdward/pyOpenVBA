"""Selecting a point or a series of a chart, replayed from live Excel (scripts/measure_chart_element_select.py,
tests/fixtures/chart_element_select.json): what Selection and ActiveChart answer afterwards, and what moves them on."""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

RECORD: dict[str, Any] = json.loads(
    (Path(__file__).parent / 'fixtures/chart_element_select.json').read_text(encoding='utf-8'))
#: The probe's last read goes on through a chart object already deleted, which is refused; its selection is tested
#: below.
DETACHED = '.Select: ch.Parent.Delete'
KEPT = [index for index, case in enumerate(RECORD['cases']) if DETACHED not in case[0]]


@pytest.fixture(scope='module')
def app() -> ExcelApplication:
    made = ExcelApplication()
    made.add_workbook()
    made.add_module(RECORD['source'], name='SelectProbe')
    return made


@pytest.mark.parametrize('index', KEPT, ids=[RECORD['cases'][index][0] for index in KEPT])
def test_each_case_as_excel_answered(app: ExcelApplication, index: int) -> None:
    assert str(app.run('SelectProbe.Probe', index, index)) == RECORD['cases'][index][1]


@pytest.mark.parametrize('element', ['ch.SeriesCollection(1)', 'ch.SeriesCollection(1).Points(2)'])
def test_deleting_the_chart_leaves_a_range_selected(element: str) -> None:
    made = ExcelApplication()
    made.add_workbook()
    made.add_module('Public Function Go() As String\n    Dim ch As Object\n    ActiveSheet.Range("B1:B3").Value = 1\n'
                    '    Set ch = ActiveSheet.ChartObjects.Add(10, 10, 200, 100).Chart\n'
                    f'    ch.SetSourceData ActiveSheet.Range("B1:B3")\n    {element}.Select\n    ch.Parent.Delete\n'
                    '    Go = TypeName(Selection) & "," & TypeName(ActiveChart) & "," & ActiveCell.Address\n'
                    'End Function\n',
                    name='Deleted')
    assert made.run('Deleted.Go') == 'Range,Nothing,$A$1'
