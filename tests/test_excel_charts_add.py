"""Charts.Add and Sheets.Add with Type xlChart, replayed from live Excel (scripts/measure_charts_add.py,
tests/fixtures/charts_add.json): what each case answers and plots, the tabs it leaves, and the parts it saves."""

from __future__ import annotations

import json
import re
import zipfile
from collections.abc import Iterator
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError
from test_excel_chart_location import plain

FIXTURE: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/charts_add.json').read_text(encoding='utf-8'))
CASES: list[list[Any]] = FIXTURE['cases']
#: The chart sheets' own parts, their drawings among them, and the list of tabs; the workbook's relationship ids and content
#: type order are the package's, and a worksheet's drawing is not what this measures.
SHEET_PARTS = re.compile(r'xl/chartsheets/(_rels/)?sheet\d+\.xml(\.rels)?|docProps/app\.xml')


def _sheet_parts(parts: dict[str, str]) -> set[str]:
    """The parts SHEET_PARTS names, and the drawings the chart sheets hold, with their relationships."""
    drawings = {f'xl/drawings/{found}' for name, text in parts.items() if name.startswith('xl/chartsheets/_rels/')
                for found in re.findall(r'Target="\.\./drawings/([^"]+)"', text)}
    drawings |= {f'xl/drawings/_rels/{name.rsplit("/", 1)[1]}.rels' for name in drawings}
    return {name for name in parts if SHEET_PARTS.fullmatch(name)} | (drawings & set(parts))
#: What the window gives a new chart sheet, which Excel works out again on opening (see _chart_sheets._NEW_SHEET).
_WINDOW = re.compile(r' zoomScale="\d+"|<xdr:ext cx="\d+" cy="\d+"/>')


@pytest.fixture(scope='module')
def app() -> ExcelApplication:
    made = ExcelApplication(sandbox=False)
    made.add_workbook()
    made.add_module(FIXTURE['source'], name='P')
    return made


@pytest.fixture(scope='module')
def folder(tmp_path_factory: pytest.TempPathFactory) -> Path:
    return tmp_path_factory.mktemp('charts_add')


@pytest.fixture(autouse=True)
def only_the_host_book(app: ExcelApplication) -> Iterator[None]:
    yield
    for book in app.application.workbooks_.books[1:]:
        book.vba_call('Close', [False])


@pytest.mark.parametrize('index', range(len(CASES)), ids=lambda index: CASES[index][0])
def test_each_case_answers_and_leaves_the_tabs_as_excel_did(index: int, app: ExcelApplication, folder: Path) -> None:
    assert str(app.run('P.Probe', str(folder), index, index)) == CASES[index][1]


@pytest.mark.parametrize('index', [index for index in range(len(CASES)) if any(name.startswith('xl/charts/chart') for name in CASES[index][2])],
                         ids=lambda index: CASES[index][0])
def test_each_chart_part_as_excel_saved_it(index: int, app: ExcelApplication, folder: Path) -> None:
    app.run('P.Probe', str(folder), index, index)
    with zipfile.ZipFile(folder / f'c{index}.xlsx') as package:
        ours = [package.read(name).decode('utf-8') for name in package.namelist() if name.startswith('xl/charts/chart')]
    theirs = [text for name, text in CASES[index][2].items() if name.startswith('xl/charts/chart')]
    assert sorted(map(plain, ours)) == sorted(map(plain, theirs))


def _same_id(text: str) -> str:
    return _WINDOW.sub('', re.sub(r'\{[0-9A-F-]{36}\}', '{id}', text))


@pytest.mark.parametrize('index', [index for index in range(len(CASES)) if any('chartsheets' in name for name in CASES[index][2])],
                         ids=lambda index: CASES[index][0])
def test_each_sheet_part_as_excel_saved_it(index: int, app: ExcelApplication, folder: Path) -> None:
    app.run('P.Probe', str(folder), index, index)
    with zipfile.ZipFile(folder / f'c{index}.xlsx') as package:
        saved = {name: package.read(name).decode('utf-8') for name in package.namelist()}
    theirs = CASES[index][2]
    assert {name: _same_id(saved[name]) for name in _sheet_parts(saved)} == {name: _same_id(theirs[name]) for name in _sheet_parts(theirs)}


LOADED: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/charts_add_loaded.json').read_text(encoding='utf-8'))
#: The cases with a series of a chart sheet read from a file deleted, which the model reports rather than does.
LOADED_REFUSED = {index for index, case in enumerate(LOADED['cases']) if 'SeriesCollection(1).Delete' in case[2]}


def _plotted_key(range_: str) -> str:
    return str(sorted({case[0] for case in LOADED['cases']}).index(range_))


@pytest.fixture(scope='module')
def loaded_app() -> ExcelApplication:
    made = ExcelApplication(sandbox=False)
    made.add_workbook()
    made.add_module(LOADED['source'], name='P')
    return made


@pytest.mark.parametrize('index', range(len(LOADED['cases'])), ids=lambda index: LOADED['cases'][index][2])
def test_beside_a_chart_sheet_read_from_a_file_as_excel_plotted(index: int, loaded_app: ExcelApplication, tmp_path: Path) -> None:
    """Placed before a chart sheet read from a file: nothing to plot until it is activated in the session, and then the cells
    its series took when it first was (scripts/measure_charts_add_loaded.py)."""
    range_, selected, _, answer = LOADED['cases'][index]
    try:
        if index in LOADED_REFUSED:
            with pytest.raises(VBAUnsupportedError):
                loaded_app.run('P.Probe', str(tmp_path / f'l{index}.xlsx'), _plotted_key(range_), selected, index)
        else:
            assert str(loaded_app.run('P.Probe', str(tmp_path / f'l{index}.xlsx'), _plotted_key(range_), selected, index)) == answer
    finally:
        for book in loaded_app.application.workbooks_.books[1:]:
            book.vba_call('Close', [False])
