"""A typed point label given a font, a turn to an area, pie or doughnut chart, and one series turned to a line, on a new chart and on
the plain loaded book, against what Excel answers and saves (scripts/measure_chart_turns_and_label_fonts.py,
tests/fixtures/chart_turns_and_label_fonts.json).

Excel writes a typed label's font in its text, whichever was set first, and a label with a font and no text in a txPr. A loaded chart
turns as a new one does: a setting the file spells as its type's default reads as unset. On the loaded book a label's font changes
in its series' labels, as its text does.

A new chart's unique ids and axis ids are random at Excel's save, so they are masked, and nothing else; the loaded book's are masked
only where the original part does not hold them."""

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

FIXTURES = Path(__file__).parent / 'fixtures'
RECORD: dict[str, Any] = json.loads((FIXTURES / 'chart_turns_and_label_fonts.json').read_text(encoding='utf-8'))
LOADED = 'loaded_chart_plain.xlsx'
_UNIQUE = re.compile(r'(uniqueId val="\{[0-9A-F]{8})-[0-9A-F-]+')
_AXIS = re.compile(r'(<c:(?:axId|crossAx) val=")\d+')


def what(case: list[Any]) -> str:
    return RECORD['edits'][case[1]][0]


def label(case: list[Any]) -> str:
    return f'{case[0]} {what(case)}'


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
        shutil.copy(FIXTURES / LOADED, folder / 'in.xlsx')
        answer = str(app.run('P.Loaded', str(folder / 'in.xlsx'), str(out), case[1]))
    with zipfile.ZipFile(out) as package:
        return answer, package.read('xl/charts/chart1.xml').decode('utf-8')


def test_every_case_was_measured() -> None:
    assert all(case[2].startswith('0 ') for case in RECORD['cases']), [label(case) for case in RECORD['cases'] if not case[2].startswith('0 ')]


@pytest.mark.parametrize('case', RECORD['cases'], ids=label)
def test_what_is_answered_is_excels_answer(case: list[Any], tmp_path: Path) -> None:
    assert run(case, tmp_path)[0] == case[2]


@pytest.mark.parametrize('case', RECORD['cases'], ids=label)
def test_the_chart_saves_as_excel_saved_it(case: list[Any], tmp_path: Path) -> None:
    assert comparable(case, run(case, tmp_path)[1]) == comparable(case, case[3]['xl/charts/chart1.xml'])
