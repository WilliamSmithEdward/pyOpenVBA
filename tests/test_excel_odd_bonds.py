"""ODDFPRICE, ODDFYIELD, ODDLPRICE and ODDLYIELD against native Excel's bits.

The functions are pyOfficeEditor's (its commit 4dc2270). These 194 cells
were measured on their own (tests/fixtures/odd_bonds/measure_native.py):
every basis, annual, semiannual and quarterly coupons, short and long odd
periods, month ends, truncated dates and errors, each double kept as its
IEEE 754 bits. All match to the bit but ODDFYIELD with a long odd first
period, where Excel's iteration takes another slope: those nine come
within 1.6e-11. A settlement on the 1900 calendar's false 29 February is
#NUM! in Excel, one of the early-1900 dates
tests/test_formula_early_1900_dates.py records.
"""
from __future__ import annotations

import hashlib
import json
import math
import struct
from pathlib import Path
from typing import Any
from zipfile import ZipFile

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.interpreter._values import VBAErrorValue

ROOT = Path(__file__).parent / 'fixtures/odd_bonds'
RAW: dict[str, Any] = json.loads((ROOT / 'native.json').read_bytes())
MANIFEST: dict[str, Any] = json.loads((ROOT / 'manifest.json').read_bytes())
ROWS = [row.split('|') for row in RAW['raw_result'].splitlines()]
CASES: list[dict[str, Any]] = RAW['cases']
#: The furthest ODDFYIELD with a long odd first period lands from Excel's answer on this batch.
LONG_FIRST_YIELD_TOLERANCE = 1.6e-11


def expected(row: list[str]) -> float | VBAErrorValue:
    # Error 2036 is #NUM!, 2015 #VALUE! and 2042 #N/A.
    if row[1] == '10':
        return VBAErrorValue(int(row[2].removeprefix('Error ')))
    assert row[1] == '5'
    return struct.unpack('>d', bytes.fromhex(row[2]))[0]


def assert_answer(case: dict[str, Any], actual: object, oracle: object) -> None:
    if case['name'].startswith('ODDFYIELD-long') and isinstance(oracle, float):
        assert isinstance(actual, (int, float))
        assert math.isclose(actual, oracle, rel_tol=0, abs_tol=LONG_FIRST_YIELD_TOLERANCE)
    else:
        assert actual == oracle


@pytest.mark.parametrize('index', [
    pytest.param(index, marks=pytest.mark.xfail(
        reason='Excel gives #NUM! for a bond settled at the start of 1900 (tests/fixtures/formula/early_1900_dates.json)',
        strict=True)) if case['name'] == 'calendar-edge-9' else index
    for index, case in enumerate(CASES)
], ids=[case['name'] for case in CASES])
def test_native_odd_coupon_value_or_error(index: int) -> None:
    app = ExcelApplication()
    app.add_workbook()
    case = CASES[index]
    app.sheet(1).set_value('A1', case['formula'])
    assert_answer(case, app.sheet(1).value('A1'), expected(ROWS[index]))


def test_native_fixture_source_input_and_complete_oracle_are_unchanged() -> None:
    assert len(CASES) == len(ROWS) == 194
    assert RAW['completed'] and RAW['closed_utc'] and RAW['status']['ok']
    assert RAW['diagnostic']['ok'] and RAW['diagnostic']['value'] == [[1.25, 'x'], [2.5, 'y']]
    assert hashlib.sha256((ROOT / 'native.json').read_bytes()).hexdigest() == MANIFEST['native_sha256']
    assert hashlib.sha256((ROOT / 'native.xlsx').read_bytes()).hexdigest() == MANIFEST['input_sha256']
    assert hashlib.sha256((ROOT / 'measure_native.py').read_bytes()).hexdigest() == MANIFEST['script_sha256']
    assert hashlib.sha256(RAW['source'].encode()).hexdigest() == RAW['source_sha256']
    with ZipFile(ROOT / 'native.xlsx') as package:
        assert not any('vbaproject' in name.casefold() for name in package.namelist())
        assert {name: hashlib.sha256(package.read(name)).hexdigest() for name in package.namelist()} == RAW['parts']


@pytest.mark.parametrize('index', [0, 1, 4, 5])
def test_native_workbook_recalculates_and_saves(index: int, tmp_path: Path) -> None:
    app = ExcelApplication.open(ROOT / 'native.xlsx', with_vba=False)
    case = CASES[index]
    cell = f'A{index + 1}'
    assert app.sheet(1).formula(cell) == case['formula']
    assert_answer(case, app.sheet(1).value(cell), expected(ROWS[index]))
    app.sheet(1).set_value(cell, case['formula'])
    output = app.save(tmp_path / 'odd-bonds.xlsx')
    reopened = ExcelApplication.open(output, with_vba=False)
    assert_answer(case, reopened.sheet(1).value(cell), expected(ROWS[index]))
