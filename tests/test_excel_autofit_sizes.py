"""Column AutoFit, Range.Text and the widening a format gives a column at every font size, measured by
scripts/measure_autofit_sizes.py.

A text is as wide as GDI makes it at the font's height in pixels, the size
times 4/3; AutoFit then adds 2 pixels to the column for every 4 the font's
digits are wider than 5, and italic reaches 57m/29 + 1/50 pixels further, m
being a fifth of the font's ascent. Range.Text keeps the same 2 for every 4
beside a text, and a pixel for each of italic's steps. Below 8pt Excel draws
both fonts in another, which reports itself unsupported.
"""
from __future__ import annotations

import json
import re
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/autofit_sizes.json').read_text(encoding='utf-8'))
PROBES: dict[str, dict[str, Any]] = {probe['name']: probe for probe in RECORD['probes']}
#: The sizes Excel draws in another font than the cell's.
SMALL = (6.0, 7.0)
#: The widening cases that differ: a number General shows, given General after its value, widens its column from
#: 14pt to a width of its own, which the model does not work out yet.
GENERAL_WIDENING = 'a number General shows widens its column when General is set on it; the model leaves it'


def _run(source: str) -> object:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(source, name='Harness')
    return app.run('Harness.Probe')


def _sizes(result: str) -> dict[str, list[str]]:
    """A fits probe's result by size: each text's width."""
    return {size: widths.split() for part in result.split('#') if part.strip()
            for size, _, widths in [part.partition(':')]}


@pytest.mark.parametrize('name', [name for name in PROBES if name.startswith('fits_')])
def test_autofit_at_every_size(name: str) -> None:
    macro = PROBES[name]['macro']
    listed = re.search(r'sizes = Array\(([^)]*)\)', macro)
    assert listed is not None
    sizes = [size for size in listed.group(1).split(', ') if float(size) not in SMALL]
    expected = {size: widths for size, widths in _sizes(PROBES[name]['result']).items() if float(size) not in SMALL}
    got = _sizes(str(_run(macro.replace(listed.group(0), f'sizes = Array({", ".join(sizes)})'))))
    assert got == expected


@pytest.mark.parametrize('size', SMALL)
def test_a_size_below_8pt_is_not_measured(size: float) -> None:
    with pytest.raises(VBAUnsupportedError, match='8 to 72pt'):
        _run(f'Function Probe() As String\nRange("A1").Value = "abc"\nRange("A1").Font.Size = {size:g}\n'
             f'Columns(1).AutoFit\nProbe = Columns(1).ColumnWidth\nEnd Function\n')


@pytest.mark.parametrize('name', [name for name in PROBES if name.startswith('narrow_')])
def test_text_at_every_width(name: str) -> None:
    expected = [line.split('\t') for line in PROBES[name]['result'].split('\n')]
    got = [line.split('\t') for line in str(_run(PROBES[name]['macro'])).split('\n')]
    assert len(got) == len(expected)
    wrong = [(theirs[0], index, one, other) for mine, theirs in zip(got, expected, strict=True)
             for index, (one, other) in enumerate(zip(mine, theirs, strict=True)) if one != other]
    assert not wrong, wrong[:5]


def _widening(result: object) -> dict[str, str]:
    return dict(part.split('=', 1) for part in str(result).split('#') if part)


WIDENING = [(name, case) for name in PROBES if name.startswith('widening_')
            for case in _widening(PROBES[name]['result'])]


@pytest.fixture(scope='module')
def widened() -> dict[str, dict[str, str]]:
    return {name: _widening(_run(PROBES[name]['macro'])) for name in PROBES if name.startswith('widening_')}


@pytest.mark.parametrize(('name', 'case'), [
    pytest.param(name, case, id=f'{name}-{case}', marks=[pytest.mark.xfail(reason=GENERAL_WIDENING, strict=True)]
                 if case.endswith('/1') and _widening(PROBES[name]['result'])[case].split('~')[0] != '64' else [])
    for name, case in WIDENING])
def test_a_format_widens_a_column_at_every_size(widened: dict[str, dict[str, str]], name: str, case: str) -> None:
    assert widened[name][case] == _widening(PROBES[name]['result'])[case]
