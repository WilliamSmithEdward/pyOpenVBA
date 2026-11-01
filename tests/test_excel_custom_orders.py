"""Custom sort orders replayed from live Excel (scripts/measure_custom_orders.py, tests/fixtures/custom_orders.json):
Sort.SortFields.Add's CustomOrder as a list's text or number, descending and matching case; Range.Sort's
OrderCustom; AddCustomList of a range mixing text and numbers; and the sortState each save writes."""
from __future__ import annotations

import json
import os
import re
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.apps.excel._sort import DESCENDING, SortObject

RECORD: dict[str, Any] = json.loads(
    (Path(__file__).parent / 'fixtures/custom_orders.json').read_text(encoding='utf-8'))
#: The probe as recorded, its files joined to the folder with this platform's separator.
SOURCE = RECORD['source'].replace('folder & "\\', f'folder & "{os.sep}')


@pytest.fixture(scope='module')
def run(tmp_path_factory: pytest.TempPathFactory) -> tuple[list[str], Path]:
    folder = tmp_path_factory.mktemp('custom_orders')
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(SOURCE, name='OrdersProbe')
    return str(app.run('OrdersProbe.Probe', str(folder))).split('@@@'), folder


@pytest.mark.parametrize('index', range(len(RECORD['cases'])))
def test_each_case_as_excel_answered(run: tuple[list[str], Path], index: int) -> None:
    assert run[0][index] == RECORD['cases'][index][3]


def _conditions(sort_state: str) -> list[dict[str, str]]:
    """A sortState's own attributes and each condition's, namespaces left out."""
    found = re.findall(r'<(?:sortState|sortCondition)\b[^>]*>', sort_state)
    return [{key: value for key, value in re.findall(r'([\w:]+)="([^"]*)"', element) if not key.startswith('xmlns')}
            for element in found]


@pytest.mark.parametrize('saved', [
    pytest.param(saved, marks=pytest.mark.xfail(strict=True, reason="Range.Sort's sortState, which the model does "
                                                                    "not write"))
    if saved.startswith('range') else saved for saved in RECORD['parts']])
def test_each_save_keeps_the_sort_state_excel_wrote(run: tuple[list[str], Path], saved: str) -> None:
    with zipfile.ZipFile(run[1] / saved) as package:
        sheet = package.read('xl/worksheets/sheet1.xml').decode('utf-8')
    found = re.search(r'<sortState\b[^>]*/>|<sortState\b.*?</sortState>', sheet, re.DOTALL)
    assert _conditions(found.group(0) if found else '') == _conditions(RECORD['parts'][saved])


def test_a_saved_custom_order_reads_back(run: tuple[list[str], Path]) -> None:
    book = ExcelApplication().open_workbook(str(run[1] / 'down.xlsx'))
    field = SortObject(book.sheets_[0]).state.fields[0]
    assert (field.order, field.key.first.address(absolute=False), field.custom) == \
        (DESCENDING, 'A1:A7', ('High', 'Medium', 'Low'))
