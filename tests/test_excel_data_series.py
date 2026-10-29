"""Range.DataSeries, replayed from live Excel (scripts/measure_data_series.py, tests/fixtures/data_series.json): the
probe's own macro runs in the model, and every cell it reads back is compared with Excel's, a Double by its bits.

The probe copies a Double's bytes with LSet of one record onto another, which the model does not implement, so its
Bits function answers a mark here and each Double's bits are read from the model's cell instead, in the same order
For Each walks them. A trend over two or more numbers is fitted by least squares, whose bits the model does not
reproduce (see _autofill), so those cases are held to reporting themselves unsupported.
"""
from __future__ import annotations

import json
import re
import struct
from pathlib import Path
from typing import Any

import pytest

from pyopenvba._a1 import parse_reference
from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

RECORD: dict[str, Any] = json.loads(
    (Path(__file__).parent / 'fixtures/data_series.json').read_text(encoding='utf-8'))
#: Cases whose trend Excel fits by least squares.
FITTED = [index for index, (setup, _, arguments, _) in enumerate(RECORD['cases'])
          if 'Trend:=True' in arguments and sum(bool(re.search(r'\.Value = [-0-9]', line)) for line in setup) >= 2]
MARK = 'BITS'


@pytest.fixture(scope='module')
def app() -> ExcelApplication:
    made = ExcelApplication()
    made.add_workbook()
    source = re.sub(r'Private Function Bits\(ByVal value As Variant\) As String.*?End Function',
                    f'Private Function Bits(ByVal value As Variant) As String\n    Bits = "{MARK}"\nEnd Function',
                    RECORD['source'], flags=re.DOTALL)
    source = source.replace('Public Function Probe() As String',
                            'Public Function One(ByVal which As Long) As String\n    One = Tried(ActiveSheet, which)\n'
                            'End Function\n\nPublic Function Probe() As String')
    made.add_module(source, name='SeriesProbe')
    return made


def _bits(app: ExcelApplication, target: str) -> list[str]:
    """The bits of each Double the probe met, in For Each's order over the target's cells."""
    sheet = app.workbook.sheets_[0]
    out: list[str] = []
    for area in parse_reference(target):
        for row in range(area.top, area.bottom + 1):
            for column in range(area.left, area.right + 1):
                cell = sheet.cells_.get((row, column))
                value = app.workbook.calculator.value_of(sheet.name, row, column) if cell is not None else None
                if isinstance(value, float):
                    out.append(struct.pack('>d', value).hex().upper())
    return out


@pytest.mark.parametrize('index', [index for index in range(len(RECORD['cases'])) if index not in FITTED],
                         ids=lambda index: f'{RECORD["cases"][index][1]} {RECORD["cases"][index][2]}')
def test_a_series_as_excel_fills_it(app: ExcelApplication, index: int) -> None:
    got = str(app.run('SeriesProbe.One', index))
    doubles = iter(_bits(app, RECORD['cases'][index][1]))
    got = re.sub(MARK, lambda _: next(doubles), got)
    assert got.split(',') == RECORD['cases'][index][3].split(',')


@pytest.mark.parametrize('index', FITTED, ids=lambda index: f'{RECORD["cases"][index][1]} {RECORD["cases"][index][2]}')
def test_a_fitted_trend_reports_itself(app: ExcelApplication, index: int) -> None:
    with pytest.raises(VBAUnsupportedError, match='Trend'):
        app.run('SeriesProbe.One', index)


def test_one_change_for_the_cells_filled() -> None:
    made = ExcelApplication()
    made.add_workbook()
    made.add_module(RECORD['sink'], name='SeriesSink', kind='class')
    made.add_module(RECORD['events_source'], name='SeriesEvents')
    assert made.run('SeriesEvents.Events') == RECORD['events']
