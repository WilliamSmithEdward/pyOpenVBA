"""One data label edit at a time on a loaded chart, against what Excel answers and saves (scripts/measure_loaded_chart_label_edits.py,
tests/fixtures/loaded_chart_label_edits.json). A case the model answers matches Excel's answer and the bytes of its chart part; one it
does not answer stays refused, so nothing half-built loses what a file holds.

A label Excel adds takes a unique id whose first 8 digits are a number and whose other 24 are new at random, unlike the part's own,
which every save keeps. Those 24 are masked where the original part does not hold them, and nothing else is."""

from __future__ import annotations

import json
import re
import shutil
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

FIXTURES = Path(__file__).parent / 'fixtures'
RECORD: dict[str, Any] = json.loads((FIXTURES / 'loaded_chart_label_edits.json').read_text(encoding='utf-8'))
_SERIES_EDITS = ('nothing', 'labels on', 'labels off', 'labels on then off', 'labels on twice', 'show category', 'show series name',
                 'show value off', 'position', 'number format', 'separator')
#: The cases the model answers, by book. The rest are refused: a chart with a style, labels the model does not write as the file
#: holds them (a turned or shadowed label), a point's label on a series whose labels have a look (Excel
#: gives it their place and the value alone), and the series' labels changing with a point's label of its own.
BUILT: dict[str, tuple[str, ...]] = {
    'loaded_chart_plain.xlsx': (*_SERIES_EDITS, 'labels on second', 'point label on', 'point label off', 'point label text',
                                'point label text second', 'two point labels', 'point labels on both', 'point text second point',
                                'point show category', 'point position'),
    'loaded_chart_dressed.xlsx': ('nothing',),
    'loaded_chart_labelled.xlsx': (*_SERIES_EDITS, 'point label text second', 'point text second point'),
}
_UNIQUE = re.compile(r'(uniqueId val="\{[0-9A-F]{8})(-[0-9A-F]{4}-4[0-9A-F]{3}-[89AB][0-9A-F]{3}-[0-9A-F]{12})')


def what(case: list[Any]) -> str:
    return RECORD['edits'][case[1]][0]


def label(case: list[Any]) -> str:
    return f'{case[0]} {what(case)}'


def answered(case: list[Any]) -> bool:
    return what(case) in BUILT[case[0]]


def masked(xml: str, original: str) -> str:
    """The part with each unique id's random end masked where the original does not hold it."""
    return _UNIQUE.sub(lambda hit: hit.group(0) if hit.group(2) in original else hit.group(1) + '-masked', xml)


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
