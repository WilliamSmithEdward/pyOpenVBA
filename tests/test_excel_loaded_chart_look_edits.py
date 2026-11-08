"""One fill or line edit at a time on a loaded chart, against what Excel answers and saves (scripts/measure_loaded_chart_look_edits.py,
tests/fixtures/loaded_chart_look_edits.json). A case the model answers matches Excel's answer and the bytes of its chart part, with the
random end of a unique id the save adds masked as tests/test_excel_loaded_chart_label_edits.py masks it; one it does not answer stays
refused."""

from __future__ import annotations

import json
import re
import shutil
import zipfile
from pathlib import Path
from typing import Any

import pytest
from test_excel_loaded_chart_label_edits import masked

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

FIXTURES = Path(__file__).parent / 'fixtures'
RECORD: dict[str, Any] = json.loads((FIXTURES / 'loaded_chart_look_edits.json').read_text(encoding='utf-8'))
_ALL = tuple(label for label, _ in RECORD['edits'])
#: The cases the model answers, by book. The rest are refused: a series whose look the model does not write as the file holds it; a
#: point's look on a series with a look, which Excel copies in; a point on a chart with a trendline, whose numbering is not known; and
#: the Interior of a series on a chart with a style.
BUILT: dict[str, tuple[str, ...]] = {
    'loaded_chart_plain.xlsx': _ALL,
    'loaded_chart_dressed.xlsx': ('nothing', 'fill colour', 'fill colour second', 'fill colour twice', 'fill hidden',
                                  'fill colour transparency', 'fill solid', 'line colour', 'line weight', 'line colour weight',
                                  'line dash', 'fill and line'),
    'loaded_chart_labelled.xlsx': _ALL,
    'loaded_chart_themed.xlsx': ('nothing',),
}


def what(case: list[Any]) -> str:
    return RECORD['edits'][case[1]][0]


def label(case: list[Any]) -> str:
    return f'{case[0]} {what(case)}'


def answered(case: list[Any]) -> bool:
    return what(case) in BUILT[case[0]]


def run(case: list[Any], folder: Path) -> tuple[str, dict[str, str]]:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(RECORD['source'], name='P')
    book, out = folder / 'in.xlsx', folder / 'out.xlsx'
    shutil.copy(FIXTURES / case[0], book)
    answer = str(app.run('P.Act', str(book), str(out), case[1]))
    with zipfile.ZipFile(out) as package:
        return answer, {name: package.read(name).decode('utf-8') for name in package.namelist() if re.fullmatch(r'xl/charts/chart\d+\.xml', name)}


def test_every_case_was_measured() -> None:
    assert all(case[2].startswith('0 ') for case in RECORD['cases']), [label(case) for case in RECORD['cases'] if not case[2].startswith('0 ')]
    assert all(set(names) <= {label for label, _ in RECORD['edits']} for names in BUILT.values())


@pytest.mark.parametrize('case', [case for case in RECORD['cases'] if not answered(case)], ids=label)
def test_an_edit_not_built_stays_refused(case: list[Any], tmp_path: Path) -> None:
    with pytest.raises(VBAUnsupportedError):
        run(case, tmp_path)


@pytest.mark.parametrize('case', [case for case in RECORD['cases'] if answered(case)], ids=label)
def test_what_is_answered_is_excels_answer(case: list[Any], tmp_path: Path) -> None:
    assert run(case, tmp_path)[0] == case[2]


@pytest.mark.parametrize('case', [case for case in RECORD['cases'] if answered(case)], ids=label)
def test_the_chart_saves_as_excel_saved_it(case: list[Any], tmp_path: Path) -> None:
    original = RECORD['originals'][case[0]]['xl/charts/chart1.xml']
    saved = {name: masked(text, original) for name, text in case[3].items() if re.fullmatch(r'xl/charts/chart\d+\.xml', name)}
    assert {name: masked(text, original) for name, text in run(case, tmp_path)[1].items()} == saved
