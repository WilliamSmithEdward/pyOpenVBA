"""Circular references and iterative calculation, measured by scripts/measure_iteration.py.

A loop is the cells a formula reads as it is worked out, back to itself.
With Application.Iteration off, a formula that closes a loop is 0, so is
any formula reading a cell of the loop, and the loop's other cells keep
what they held; none is worked out again until the loop is broken, and
Worksheet.CircularReference names a cell of the sheet's loops, which of
them changing run to run in Excel. With it on, every calculation works
the loop's cells out again, and those of what reads them, a sheet at a
time in the order the sheets were made and row by row, each reading what
the others hold, until no number changes by MaxChange or more or
MaxIterations passes are done; a cell starts at 0, and text and logicals
change nothing.
"""
from __future__ import annotations

import json
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBARuntimeError

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/iteration.json').read_text(encoding='utf-8'))
PROBES: dict[str, dict[str, Any]] = {probe['name']: probe for probe in RECORD['probes']}
#: Formulas naming their own cell, measured by scripts/measure_self_reference.py: a loop is closed only through a cell
#: a formula reads as it is worked out, which ROWS, ROW, CELL("address"), FORMULATEXT, an IF's branch not taken,
#: IFERROR's fallback when nothing fails and a MATCH or VLOOKUP stopping at its match before the cell do not.
SELF_REFERENCES: list[dict[str, Any]] = json.loads(
    (Path(__file__).parent / 'fixtures/self_reference.json').read_text(encoding='utf-8'))['probes']


#: Probes the model does not answer as Excel does, and why.
GAPS: dict[str, str] = {
    'off_stale_member': 'with iteration off, A1 left to work out again by the edit of A2 and read by nothing before '
                        'A3 closed the loop keeps what it was before that edit, 1, where Excel had worked it out '
                        'after the edit, 2: the model works a formula out when it is read',
}


@pytest.mark.parametrize('name', [
    pytest.param(name, marks=[pytest.mark.xfail(reason=GAPS[name], strict=True)] if name in GAPS else [])
    for name in PROBES])
def test_a_circular_reference(name: str) -> None:
    probe = PROBES[name]
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(probe['macro'], name='Harness')
    if 'error' in probe:
        with pytest.raises(VBARuntimeError) as raised:
            app.run('Harness.Probe')
        assert raised.value.number == probe['error']
    else:
        # Of two cells of one loop on a sheet, CircularReference names either, run to run: any answer Excel gave in
        # the repeated runs is one.
        assert app.run('Harness.Probe') in {probe['result'], *RECORD['repeated'].get(name, [])}


@pytest.mark.parametrize('probe', SELF_REFERENCES, ids=lambda probe: probe['name'])
def test_a_formula_naming_its_own_cell(probe: dict[str, Any]) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(probe['macro'], name='Harness')
    assert app.run('Harness.Probe') == probe['result']


@pytest.mark.parametrize('name', list(RECORD['file']['settings']))
def test_the_settings_a_workbook_keeps(tmp_path: Path, name: str) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(f"Sub Setting()\n{RECORD['file']['settings'][name]}\nEnd Sub\n", name='Harness')
    app.run('Harness.Setting')
    saved = app.save(tmp_path / f'{name}.xlsx')
    with zipfile.ZipFile(saved) as package:
        workbook = package.read('xl/workbook.xml').decode('utf-8')
    expected = RECORD['file']['calcPr'][name].replace(' calcId="191029"', '')
    found = workbook[workbook.index('<calcPr'):]
    found = found[:found.index('>') + 1]
    for attribute in ('iterate', 'iterateCount', 'iterateDelta'):
        assert (f' {attribute}="' in found) == (f' {attribute}="' in expected), (attribute, found, expected)
    for piece in expected.removeprefix('<calcPr').removesuffix('/>').split():
        assert piece in found, (piece, found)


def _settings(app: ExcelApplication) -> str:
    app.add_module('Function Probe() As String\nProbe = Application.Iteration & "|" & Application.MaxIterations '
                   '& "|" & Application.MaxChange\nEnd Function\n', name='Reading')
    return str(app.run('Reading.Probe'))


def test_the_first_workbook_brings_its_settings_and_a_later_one_does_not(tmp_path: Path) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(f"Sub Setting()\n{RECORD['file']['settings']['iterating']}\nEnd Sub\n", name='Harness')
    app.run('Harness.Setting')
    saved = app.save(tmp_path / 'iterating.xlsx')
    opened = RECORD['file']['opened'].split('|')
    # Opened in an Excel of its own, it brings them; opened beside another workbook, it leaves the application's.
    assert _settings(ExcelApplication.open(saved)) == '|'.join(opened[6:9])
    later = ExcelApplication()
    later.add_workbook()
    later.open_workbook(saved)
    assert _settings(later) == '|'.join(opened[:3])
