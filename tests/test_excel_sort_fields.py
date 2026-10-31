"""Sort fields replayed from live Excel (scripts/measure_sort_fields.py, tests/fixtures/sort_fields.json): SortField's
Priority, Delete, ModifyKey, CustomOrder, SubField, Parent and Application, SortFields' Parent, items and For Each,
and the sortState each save writes after them."""
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
    (Path(__file__).parent / 'fixtures/sort_fields.json').read_text(encoding='utf-8'))
#: The probe as recorded, its files joined to the folder with this platform's separator.
SOURCE = RECORD['source'].replace('folder & "\\', f'folder & "{os.sep}')


@pytest.fixture(scope='module')
def run(tmp_path_factory: pytest.TempPathFactory) -> tuple[list[str], Path]:
    folder = tmp_path_factory.mktemp('sort_fields')
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(SOURCE, name='FieldsProbe')
    return str(app.run('FieldsProbe.Probe', str(folder))).split('@@@'), folder


@pytest.mark.parametrize('index', range(len(RECORD['cases'])))
def test_each_case_as_excel_answered(run: tuple[list[str], Path], index: int) -> None:
    assert run[0][index] == RECORD['cases'][index][3]


def _conditions(sort_state: str) -> list[dict[str, str]]:
    """A sortState's own attributes and each condition's in order, namespaces left out."""
    found = re.findall(r'<(?:sortState|sortCondition)\b[^>]*>', sort_state)
    return [{key: value for key, value in re.findall(r'([\w:]+)="([^"]*)"', element) if not key.startswith('xmlns')}
            for element in found]


@pytest.mark.parametrize('saved', list(RECORD['parts']))
def test_each_save_keeps_the_sort_state_excel_wrote(run: tuple[list[str], Path], saved: str) -> None:
    with zipfile.ZipFile(run[1] / saved) as package:
        sheet = package.read('xl/worksheets/sheet1.xml').decode('utf-8')
    found = re.search(r'<sortState\b[^>]*/>|<sortState\b.*?</sortState>', sheet, re.DOTALL)
    assert _conditions(found.group(0) if found else '') == _conditions(RECORD['parts'][saved])


def test_a_sub_field_is_not_saved(tmp_path: Path) -> None:
    """Excel's file for a field sorting by a data type's field is not measured."""
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('Sub Save(ByVal path As String)\n    Dim ws As Worksheet\n    Set ws = ActiveSheet\n'
                   '    ws.Sort.SortFields.Add ws.Range("A1:A3")\n    ws.Sort.SetRange ws.Range("A1:A3")\n'
                   '    ws.Sort.SortFields(1).SubField = "Population"\n'
                   '    ActiveWorkbook.SaveAs path, xlOpenXMLWorkbook\nEnd Sub\n', name='Saving')
    with pytest.raises(VBAUnsupportedError, match='SubField'):
        app.run('Saving.Save', str(tmp_path / 'sub.xlsx'))
