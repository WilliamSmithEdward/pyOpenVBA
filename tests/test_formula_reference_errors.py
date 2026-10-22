"""Errors and arrays given to a function's range-only parameter, measured by scripts/measure_reference_errors.py.

An error there is handed on, whether the formula holds a #REF! or LET
hands over an error value; ISREF alone says FALSE. An array there is
taken one item at a time, every item #VALUE!, as pyOfficeEditor
measured. Where LET hands that array over in a formula Formula2 writes,
Excel keeps the formula as a dynamic-array one, which reads back with
Formula as @LET(...): the model keeps it a legacy formula, so it does
not spill. Those cases are listed in GAPS and must still differ.
"""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

PROBES: list[dict[str, Any]] = json.loads(
    (Path(__file__).parent / 'fixtures/formula/reference_errors.json').read_text(encoding='utf-8'))['probes']
#: The items of a probe's answer the model does not give yet: a LET-bound array in a range-only parameter of a
#: formula Formula2 writes, which Excel makes a dynamic-array formula. In the array probe these are COUNTIF,
#: COUNTIFS, SUMIF, SUMIFS, AVERAGEIF, MAXIFS, COUNTBLANK, RANK, AREAS, SUBTOTAL and DSUM.
GAPS: dict[str, frozenset[int]] = {'array': frozenset({4, 5, 6, 7, 8, 9, 10, 11, 14, 17, 18}),
                                   'readback': frozenset({0, 1, 2, 3, 8})}


@pytest.mark.parametrize('probe', PROBES, ids=lambda probe: probe['name'])
def test_range_only_parameters(probe: dict[str, Any]) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(probe['macro'], name='Harness')
    got = str(app.run('Harness.Probe')).split('|')
    want = probe['result'].split('|')
    assert len(got) == len(want)
    gaps = GAPS.get(probe['name'], frozenset())
    for index, (excel, model) in enumerate(zip(want, got, strict=True)):
        if index in gaps:
            assert model != excel, f'item {index} now matches Excel: take it out of GAPS'
        else:
            assert model == excel, f'item {index}'
