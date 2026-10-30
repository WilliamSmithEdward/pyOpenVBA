"""Issue #252 follow-ups replayed from en-US Excel (scripts/measure_text_compare_names.py,
tests/fixtures/text_compare_names.json): which text comparisons ignore kana type and width, and how a table takes a
name through ListObject.Name and DisplayName, with what a save keeps."""
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
    (Path(__file__).parent / 'fixtures/text_compare_names.json').read_text(encoding='utf-8'))


def test_comparisons_fold_and_searches_do_not() -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(RECORD['compare_source'], name='TextCompare')
    got = str(app.run('TextCompare.Compare')).split('|')
    assert dict(zip(RECORD['compare'], got, strict=True)) == RECORD['compare']


@pytest.fixture(scope='module')
def named(tmp_path_factory: pytest.TempPathFactory) -> tuple[list[str], Path]:
    folder = tmp_path_factory.mktemp('names')
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(RECORD['names_source'].replace('folder & "\\" & saved', f'folder & "{os.sep}" & saved'),
                   name='TableNames')
    return str(app.run('TableNames.Names', str(folder))).split('@@@'), folder


@pytest.mark.parametrize('index', range(len(RECORD['names'])),
                         ids=[f'{which}={wanted}' for which, wanted, _ in RECORD['names']])
def test_a_table_takes_a_name_as_excel_does(named: tuple[list[str], Path], index: int) -> None:
    assert named[0][index] == RECORD['names'][index][2]
    saved = f'name{index}.xlsx'
    if saved in RECORD['parts']:
        with zipfile.ZipFile(named[1] / saved) as package:
            head = re.search(r'<table\b[^>]*>', package.read('xl/tables/table1.xml').decode('utf-8'))
        assert head is not None
        assert re.findall(r'\s((?:name|displayName)="[^"]*")', head.group(0)) == RECORD['parts'][saved]
