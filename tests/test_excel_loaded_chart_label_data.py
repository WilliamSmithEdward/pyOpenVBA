"""A loaded chart's data labels when its data changes, against what Excel answers and saves
(scripts/measure_loaded_chart_label_data.py, tests/fixtures/loaded_chart_label_data.json), on the labelled book and on the plain
book with the same labels set in the session.

A series in the same place with the same cells keeps its labels across SetSourceData, PlotBy, Values and Formula; one whose cells
change loses them, its points' too. Excel kept them with a literal array in the one case measured, and gave one series its labels
back when a second SetSourceData returned it to its cells, which no rule here explains: both stay refused."""

from __future__ import annotations

import json
import shutil
import zipfile
from pathlib import Path
from typing import Any

import pytest
from test_excel_loaded_chart_label_edits import masked

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

FIXTURES = Path(__file__).parent / 'fixtures'
RECORD: dict[str, Any] = json.loads((FIXTURES / 'loaded_chart_label_data.json').read_text(encoding='utf-8'))
LABELLED = 'loaded_chart_labelled.xlsx'
#: Refused on each book: the two cases no rule explains, and Chart.PlotBy, which the model does not set.
REFUSED = ('source columns and back', 'values array', 'plot by columns')


def what(case: list[Any]) -> str:
    return RECORD['edits'][case[1]][0]


def label(case: list[Any]) -> str:
    return f'{case[0]} {what(case)}'


def run(case: list[Any], folder: Path) -> tuple[str, str]:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(RECORD['source'], name='P')
    book, out = folder / 'in.xlsx', folder / 'out.xlsx'
    shutil.copy(FIXTURES / case[0], book)
    answer = str(app.run('P.Act', str(book), str(out), case[1], '' if case[0] == LABELLED else 'labels'))
    with zipfile.ZipFile(out) as package:
        return answer, package.read('xl/charts/chart1.xml').decode('utf-8')


def test_every_case_was_measured() -> None:
    assert all(case[2].startswith('0 ') for case in RECORD['cases']), [label(case) for case in RECORD['cases'] if not case[2].startswith('0 ')]


@pytest.mark.parametrize('case', [case for case in RECORD['cases'] if what(case) in REFUSED], ids=label)
def test_an_edit_not_built_stays_refused(case: list[Any], tmp_path: Path) -> None:
    with pytest.raises(VBAUnsupportedError):
        run(case, tmp_path)


@pytest.mark.parametrize('case', [case for case in RECORD['cases'] if what(case) not in REFUSED], ids=label)
def test_what_is_answered_is_excels_answer(case: list[Any], tmp_path: Path) -> None:
    assert run(case, tmp_path)[0] == case[2]


@pytest.mark.parametrize('case', [case for case in RECORD['cases'] if what(case) not in REFUSED], ids=label)
def test_the_chart_saves_as_excel_saved_it(case: list[Any], tmp_path: Path) -> None:
    original = RECORD['originals'][case[0]]['xl/charts/chart1.xml']
    assert masked(run(case, tmp_path)[1], original) == masked(case[3]['xl/charts/chart1.xml'], original)
