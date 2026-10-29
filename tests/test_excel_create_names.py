"""Range.CreateNames, replayed from live Excel (scripts/measure_create_names.py, tests/fixtures/create_names.json):
the measuring macros run in the model in the order Excel ran them, every side given or guessed, labels of every
kind and the names they clean into.

Excel lists a hidden name, _xlfn.SINGLE, the model does not make (see tests/test_excel_workbook_links.py); the name
lists are compared without it.
"""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.apps.excel._create_names import name_of

RECORD: dict[str, Any] = json.loads(
    (Path(__file__).parent / 'fixtures/create_names.json').read_text(encoding='utf-8'))
SINGLE = '|_xlfn.SINGLE==#NAME?'


@pytest.fixture(scope='module')
def app() -> ExcelApplication:
    made = ExcelApplication()
    made.add_workbook()
    made.add_module(RECORD['source'], name='NamesProbe')
    return made


def _excels(answer: str) -> list[str]:
    return answer.replace(SINGLE, '').split('|')


@pytest.mark.parametrize('sides', list(RECORD['sides']))
def test_each_side_given(app: ExcelApplication, sides: str) -> None:
    got = app.run('NamesProbe.Sides', *(flag == '1' for flag in sides))
    assert str(got).split('|') == _excels(RECORD['sides'][sides])


@pytest.mark.parametrize('index', range(len(RECORD['guess_cases'])),
                         ids=[f'{layout}-{target}-{arguments}' for layout, target, arguments, _ in RECORD['guess_cases']])
def test_sides_guessed(app: ExcelApplication, index: int) -> None:
    assert str(app.run(f'NamesProbe.Guess{index}')).split('|') == _excels(RECORD['guess_cases'][index][3])


def test_labels_of_other_kinds(app: ExcelApplication) -> None:
    """A formula's text and dates as they show are labels, numbers, errors and blank text none; a name the
    workbook has, or one made twice, takes the later spelling; a label is cut to 255 before it is cleaned."""
    assert str(app.run('NamesProbe.Kinds')).split('|') == _excels(RECORD['kinds'])


def test_labels_cleaned_into_names(app: ExcelApplication) -> None:
    assert str(app.run('NamesProbe.Cleaning')).split('|') == _excels(RECORD['cleaning'])


@pytest.mark.parametrize(('label', 'name'), [('A1', 'A1_'), ('R1C1', '_R1C1'), ('3rd', '_3rd'), ('  lead', 'lead'),
                                              ('Total %', 'Total'), ('a  b', 'a__b'), ('%', None)])
def test_a_label_cleaned(label: str, name: str | None) -> None:
    assert name_of(label) == name
