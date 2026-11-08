"""Every loaded-chart edit measured on the plain book, on a chart with a style: one made with Shapes.AddChart2(201), and the
same chart saved and opened again, against what Excel answers and saves (scripts/measure_loaded_chart_styled_edits.py,
tests/fixtures/loaded_chart_styled_edits.json).

Excel's save of the loaded chart is its save of the new one for every edit it took on both, but a Series.Formula, so the
model's new styled chart gives the loaded one's bytes. A new chart's unique ids and axis ids are random at Excel's save, so
they are masked, and nothing else; the loaded book's are masked only where the original part does not hold them."""

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
RECORD: dict[str, Any] = json.loads((FIXTURES / 'loaded_chart_styled_edits.json').read_text(encoding='utf-8'))
BOOK = 'loaded_chart_styled.xlsx'
#: Refused on the loaded book: a title of several lines, which the model writes otherwise on a new chart too; a Series.Formula
#: changing categories; the Interior of a styled chart's element, which Excel reads as the element's own; the series' labels
#: going with a point's label of its own; and Chart.PlotBy.
REFUSED_LOADED = ('element series formula', 'element title text two lines', 'element title text cr', 'look interior colour',
                  'labels point label then labels off', 'structure plot by columns')
#: Refused on the new chart too.
REFUSED_NEW = ('labels point label then labels off', 'structure plot by columns')
#: What the model answers on the new chart otherwise than Excel: a title of several lines or none. Held so that a fix moves them.
GAPS_NEW = ('element title text two lines', 'element title text cr', 'element title text empty')
_UNIQUE = re.compile(r'(uniqueId val="\{[0-9A-F]{8})-[0-9A-F-]+')
_AXIS = re.compile(r'(<c:(?:axId|crossAx) val=")\d+')


def what(case: list[Any]) -> str:
    return RECORD['edits'][case[1]][0]


def label(case: list[Any]) -> str:
    return f'{case[0]} {what(case)}'


def excel_answered(case: list[Any]) -> bool:
    # Excel itself failed twice on a new chart's Series.Formula, killing the session.
    return 'xl/charts/chart1.xml' in case[3]


def refused(case: list[Any]) -> bool:
    return what(case) in (REFUSED_NEW if case[0] == 'new' else REFUSED_LOADED)


def answered(case: list[Any]) -> bool:
    return excel_answered(case) and not refused(case) and not (case[0] == 'new' and what(case) in GAPS_NEW)


def comparable(case: list[Any], xml: str) -> str:
    if case[0] == 'new':
        return _AXIS.sub(r'\1', _UNIQUE.sub(r'\1', xml))
    return masked(xml, RECORD['original']['xl/charts/chart1.xml'])


def run(case: list[Any], folder: Path) -> tuple[str, str]:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(RECORD['source'], name='P')
    out = folder / 'out.xlsx'
    if case[0] == 'new':
        answer = str(app.run('P.Fresh', str(out), case[1]))
    else:
        shutil.copy(FIXTURES / BOOK, folder / 'in.xlsx')
        answer = str(app.run('P.Loaded', str(folder / 'in.xlsx'), str(out), case[1]))
    with zipfile.ZipFile(out) as package:
        return answer, package.read('xl/charts/chart1.xml').decode('utf-8')


def test_the_book_is_the_one_measured() -> None:
    with zipfile.ZipFile(FIXTURES / BOOK) as package:
        assert package.read('xl/charts/chart1.xml').decode('utf-8') == RECORD['original']['xl/charts/chart1.xml']


@pytest.mark.parametrize('case', [case for case in RECORD['cases'] if refused(case)], ids=label)
def test_an_edit_not_built_stays_refused(case: list[Any], tmp_path: Path) -> None:
    with pytest.raises(VBAUnsupportedError):
        run(case, tmp_path)


@pytest.mark.parametrize('case', [case for case in RECORD['cases'] if answered(case)], ids=label)
def test_what_is_answered_is_excels_answer(case: list[Any], tmp_path: Path) -> None:
    assert run(case, tmp_path)[0] == case[2]


@pytest.mark.parametrize('case', [case for case in RECORD['cases'] if answered(case)], ids=label)
def test_the_chart_saves_as_excel_saved_it(case: list[Any], tmp_path: Path) -> None:
    assert comparable(case, run(case, tmp_path)[1]) == comparable(case, case[3]['xl/charts/chart1.xml'])


@pytest.mark.parametrize('case', [case for case in RECORD['cases'] if case[0] == 'new' and what(case) in GAPS_NEW], ids=label)
def test_a_known_gap_on_a_new_chart_still_differs(case: list[Any], tmp_path: Path) -> None:
    answer, xml = run(case, tmp_path)
    assert (answer, comparable(case, xml)) != (case[2], comparable(case, case[3]['xl/charts/chart1.xml'])), 'fixed: take it off GAPS_NEW'
