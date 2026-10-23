"""Range.Text in a column too narrow for what a cell shows, measured by scripts/measure_narrow_text.py.

A number in General loses decimals, then goes to scientific notation,
then shows 0 where it rounds to that, as its column narrows; any other
number, a date, a logical or an error fills the column with # once it no
longer fits. Text shows whole. What a cell has to show its text in is its
column's width but 5 pixels, an indent level's 9 and italic's reach; a
merged cell's is its columns' and a cell that shrinks its text shows it all.
A * fill repeats its character across what is left, to 255 characters.
Setting a number format widens a column wherever AutoFit would.
"""
from __future__ import annotations

import json
import re
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/narrow_text.json').read_text(encoding='utf-8'))
PROBES: dict[str, dict[str, Any]] = {probe['name']: probe for probe in RECORD['probes']}
#: What the model shows otherwise than Excel did: (Excel's text, the model's first character there) by why, and how
#: many cells of one probe may show it.
GAPS = {
    # -1234.5678 through *-General shows 0 where a digit fits and its sign does not, which no rounding of it is.
    ('0', '#'): ('a negative number General shows beside a fill shows 0 in the five pixels a digit alone fits', 5),
}
#: The probes that set something else on a cell, whose Text the model reports unsupported rather than guess.
UNSUPPORTED = {'hidden_column': 'hidden column'}
#: The probes of what widens a column, and the saved file after it does.
WIDENING = 'widening'


def _run(source: str) -> object:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(source, name='Harness')
    return app.run('Harness.Probe')


@pytest.mark.parametrize('name', [name for name, probe in PROBES.items() if '\t' in str(probe['result'])])
def test_text_at_every_width(name: str) -> None:
    expected = [line.split('\t') for line in PROBES[name]['result'].split('\n')]
    got = [line.split('\t') for line in str(_run(PROBES[name]['macro'])).split('\n')]
    assert len(got) == len(expected)
    wrong: list[tuple[str, int, str, str]] = []
    known: dict[tuple[str, str], int] = {}
    for mine, theirs in zip(got, expected, strict=True):
        for index, (one, other) in enumerate(zip(mine, theirs, strict=True)):
            if one == other:
                continue
            if (other, one[:1]) in GAPS:
                known[other, one[:1]] = known.get((other, one[:1]), 0) + 1
                continue
            wrong.append((theirs[0], index, one, other))
    assert not wrong, wrong[:5]
    # Each gap is as many widths as it was measured at, no more.
    assert all(count <= GAPS[key][1] for key, count in known.items()), known


@pytest.mark.parametrize('name', [name for name, probe in PROBES.items()
                                  if '\t' not in str(probe['result']) and name != WIDENING])
def test_text_with_something_else_set(name: str) -> None:
    if name in UNSUPPORTED:
        with pytest.raises(VBAUnsupportedError, match=UNSUPPORTED[name].split()[-1]):
            _run(PROBES[name]['macro'])
        return
    assert _run(PROBES[name]['macro']) == PROBES[name]['result']


def _cases(probe: dict[str, Any], result: object) -> dict[str, str]:
    """A widening probe's result case by case: each case's column width in pixels and what its cell shows."""
    names: list[str] = re.findall(r'"#(\w+)="', probe['macro'])
    text = '#' + str(result)
    out: dict[str, str] = {}
    for name, following in zip(names, [*names[1:], ''], strict=True):
        start = text.index(f'#{name}=') + len(name) + 2
        out[name] = text[start:text.index(f'#{following}=', start)] if following else text[start:]
    return out


@pytest.fixture(scope='module')
def widened() -> dict[str, str]:
    return _cases(PROBES[WIDENING], _run(PROBES[WIDENING]['macro']))


@pytest.mark.parametrize('case', list(_cases(PROBES[WIDENING], PROBES[WIDENING]['result'])))
def test_excel_widens_a_column_for_a_format(widened: dict[str, str], case: str) -> None:
    assert widened[case] == _cases(PROBES[WIDENING], PROBES[WIDENING]['result'])[case]


def test_a_widened_column_is_saved_as_autofit_saves_one(tmp_path: Path) -> None:
    path = tmp_path / 'widened.xlsx'
    assert _run(RECORD['file']['macro'].replace('{path}', str(path).replace('"', '""'))) == 'saved'
    with zipfile.ZipFile(path) as package:
        sheet = package.read('xl/worksheets/sheet1.xml').decode('utf-8')
    found = re.search(r'<cols>.*?</cols>', sheet)
    assert (found.group(0) if found else '') == RECORD['file']['cols']
