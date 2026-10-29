"""Which names Names.Add takes, as live Excel took them (scripts/measure_name_validity.py): 324 candidates, those
that start like an R1C1 reference with every ending among them."""
from __future__ import annotations

import json
from pathlib import Path

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBARuntimeError

TAKEN: dict[str, bool] = json.loads(
    (Path(__file__).parent / 'fixtures/name_validity.json').read_text(encoding='utf-8'))['taken']


@pytest.fixture(scope='module')
def app() -> ExcelApplication:
    made = ExcelApplication()
    made.add_workbook()
    made.add_module('Sub Take(ByVal nm As String)\nThisWorkbook.Names.Add Name:=nm, RefersTo:="=1"\n'
                    'ThisWorkbook.Names(nm).Delete\nEnd Sub\n', name='Naming')
    return made


@pytest.mark.parametrize('name', list(TAKEN))
def test_names_add_takes_what_excel_takes(app: ExcelApplication, name: str) -> None:
    if TAKEN[name]:
        app.run('Naming.Take', name)
    else:
        with pytest.raises(VBARuntimeError) as raised:
            app.run('Naming.Take', name)
        assert raised.value.number == 1004
