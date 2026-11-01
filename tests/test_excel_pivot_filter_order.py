"""PivotFilter.Order and PivotItems.Add, replayed from live Excel (scripts/measure_pivot_filter_order.py,
tests/fixtures/pivot_filter_order.json): every filter with its type and Order after each case, the errors, each
saved filter, and the Region field's items as the table and the cache save them."""
from __future__ import annotations

import json
import os
import re
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

RECORD: dict[str, Any] = json.loads(
    (Path(__file__).parent / 'fixtures/pivot_filter_order.json').read_text(encoding='utf-8'))
#: The probe as recorded, its files joined to the folder with this platform's separator.
SOURCE = RECORD['source'].replace('folder & "\\', f'folder & "{os.sep}')


def saved_filters(path: Path) -> list[str]:
    with zipfile.ZipFile(path) as package:
        names = package.namelist()
        table = next(one for one in names if re.fullmatch(r'xl/pivotTables/pivotTable\d+\.xml', one))
        xml = package.read(table).decode('utf-8')
        cache = package.read(next(one for one in names if re.fullmatch(
            r'xl/pivotCache/pivotCacheDefinition\d+\.xml', one))).decode('utf-8')
    found = [' '.join(re.findall(r'\b(?:fld|type|evalOrder|id)="[^"]*"', one))
             for one in re.findall(r'<filter\b[^>]*>', xml)]
    region = re.search(r'<pivotField\b[^>]*>(.*?)</pivotField>', xml, re.S)
    shared = re.search(r'<cacheField name="Region"[^>]*>(.*?)</cacheField>', cache, re.S)
    return found + [region.group(1) if region else '', shared.group(1) if shared else '']


@pytest.fixture(scope='module')
def run(tmp_path_factory: pytest.TempPathFactory) -> tuple[list[str], Path]:
    folder = tmp_path_factory.mktemp('pivot_filter_order')
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(SOURCE, name='FilterOrderProbe')
    return str(app.run('FilterOrderProbe.Probe', str(folder))).split('@@@'), folder


@pytest.mark.parametrize('index', range(len(RECORD['cases'])))
def test_each_case_as_excel_answered(run: tuple[list[str], Path], index: int) -> None:
    assert run[0][index] == RECORD['cases'][index][1]


@pytest.mark.parametrize('index', range(len(RECORD['cases'])))
def test_each_definition_as_excel_saved_it(run: tuple[list[str], Path], index: int) -> None:
    assert saved_filters(run[1] / f'c{index}.xlsx') == RECORD['cases'][index][2]
