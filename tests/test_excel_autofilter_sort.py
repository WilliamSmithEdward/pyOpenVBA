"""AutoFilter.Sort replayed from live Excel (scripts/measure_autofilter_sort.py, tests/fixtures/autofilter_sort.json):
the filter's own sort settings, apart from the sheet's; its fixed range and header; what removing the filter leaves;
and the autoFilter element each save writes before and after Apply."""
from __future__ import annotations

import json
import os
import re
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

RECORD: dict[str, Any] = json.loads(
    (Path(__file__).parent / 'fixtures/autofilter_sort.json').read_text(encoding='utf-8'))
#: The probe as recorded, its files joined to the folder with this platform's separator.
SOURCE = RECORD['source'].replace('folder & "\\', f'folder & "{os.sep}')


@pytest.fixture(scope='module')
def run(tmp_path_factory: pytest.TempPathFactory) -> tuple[list[str], Path]:
    folder = tmp_path_factory.mktemp('autofilter_sort')
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(SOURCE, name='FilterSortProbe')
    return str(app.run('FilterSortProbe.Probe', str(folder))).split('@@@'), folder


@pytest.mark.parametrize('index', range(len(RECORD['cases'])))
def test_each_case_as_excel_answered(run: tuple[list[str], Path], index: int) -> None:
    assert run[0][index] == RECORD['cases'][index][3]


def _elements(markup: str) -> list[tuple[str, dict[str, str]]]:
    """Each element's name and attributes in order, namespaces and the filter's revision id left out."""
    return [(name, {key: value for key, value in re.findall(r'([\w:]+)="([^"]*)"', attributes)
                    if not key.startswith('xmlns') and key != 'xr:uid'})
            for name, attributes in re.findall(r'<(\w+)\b([^>]*)>', markup)]


@pytest.mark.parametrize('saved', list(RECORD['parts']))
def test_each_save_writes_the_filter_excel_wrote(run: tuple[list[str], Path], saved: str) -> None:
    """The sort settings go in the autoFilter once applied, and none in the sheet's own sortState."""
    with zipfile.ZipFile(run[1] / saved) as package:
        sheet = package.read('xl/worksheets/sheet1.xml').decode('utf-8')
    found = re.search(r'<autoFilter\b[^>]*/>|<autoFilter\b.*?</autoFilter>', sheet, re.DOTALL)
    outside = re.search(r'<sortState\b[^>]*/>|<sortState\b.*?</sortState>',
                        re.sub(r'<autoFilter\b.*?</autoFilter>', '', sheet, flags=re.DOTALL), re.DOTALL)
    assert _elements(found.group(0) if found else '') == _elements(RECORD['parts'][saved][0])
    assert _elements(outside.group(0) if outside else '') == _elements(RECORD['parts'][saved][1])


def test_a_saved_filter_sort_is_not_read(run: tuple[list[str], Path]) -> None:
    """Excel's reading of the sort settings a file keeps with its AutoFilter is not measured."""
    app = ExcelApplication.open(run[1] / 'down.xlsx')
    app.add_module('Function P() As Long\n    P = ActiveSheet.AutoFilter.Sort.SortFields.Count\nEnd Function\n',
                   name='Reading')
    with pytest.raises(VBAUnsupportedError, match='AutoFilter'):
        app.run('Reading.P')
