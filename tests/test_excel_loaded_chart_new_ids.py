"""The unique ids Excel gives what it adds to a loaded chart, a point's label or fill or a series, on books with one or two
trendlines and with a manual layout, against what Excel answers and saves (scripts/measure_loaded_chart_new_ids.py,
tests/fixtures/loaded_chart_new_ids.json).

They are numbered from the chart's count of trendlines; a manual layout counts for nothing."""

from __future__ import annotations

import json
import shutil
import zipfile
from pathlib import Path
from typing import Any

import pytest
from test_excel_loaded_chart_label_edits import masked

from pyopenvba.apps.excel import ExcelApplication

FIXTURES = Path(__file__).parent / 'fixtures'
RECORD: dict[str, Any] = json.loads((FIXTURES / 'loaded_chart_new_ids.json').read_text(encoding='utf-8'))


def label(case: list[Any]) -> str:
    return f'{case[0]} {RECORD["edits"][case[1]][0]}'


def run(case: list[Any], folder: Path) -> tuple[str, str]:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(RECORD['source'], name='P')
    book, out = folder / 'in.xlsx', folder / 'out.xlsx'
    shutil.copy(FIXTURES / case[0], book)
    answer = str(app.run('P.Act', str(book), str(out), case[1]))
    with zipfile.ZipFile(out) as package:
        return answer, package.read('xl/charts/chart1.xml').decode('utf-8')


def test_every_case_was_measured() -> None:
    assert all(case[2].startswith('0 ') for case in RECORD['cases']), [label(case) for case in RECORD['cases'] if not case[2].startswith('0 ')]


def test_the_books_are_the_ones_measured() -> None:
    for name, part in RECORD['originals'].items():
        with zipfile.ZipFile(FIXTURES / name) as package:
            assert package.read('xl/charts/chart1.xml').decode('utf-8') == part, name


@pytest.mark.parametrize('case', RECORD['cases'], ids=label)
def test_what_is_answered_is_excels_answer(case: list[Any], tmp_path: Path) -> None:
    assert run(case, tmp_path)[0] == case[2]


@pytest.mark.parametrize('case', RECORD['cases'], ids=label)
def test_the_chart_saves_as_excel_saved_it(case: list[Any], tmp_path: Path) -> None:
    original = RECORD['originals'][case[0]]
    assert masked(run(case, tmp_path)[1], original) == masked(case[3], original)
