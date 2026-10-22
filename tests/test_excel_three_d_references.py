"""Formulas given a 3D reference, ThreeA:ThreeB!A1, measured by scripts/measure_three_d_references.py.

A 3D reference is not a value: alone, with an operator, behind @ or handed
back by IF it is #REF!, and ISREF calls it no reference. The statistics
functions read every sheet of it (SUM, COUNT, LARGE, MEDIAN, SUMSQ, NPV and
the others), as do AND, OR, XOR, CONCAT, TEXTJOIN, SHEET, SHEETS, VSTACK,
HSTACK, TOCOL and TOROW, each sheet's area in turn. The functions that take
an array, SUMPRODUCT, MMULT, SORT, CORREL and their kind, give #REF!; every
other function given one for a range gives #VALUE!, ROWS, INDEX, COUNTIF and
SUBTOTAL among them. Excel writes it back first sheet to last, ThreeB:ThreeA
as ThreeA:ThreeB and ThreeA:ThreeA as ThreeA. Two spellings read otherwise:
the first sheet named like a cell, Q1:Q4!A1, is the cell Q1 to Q4!A1, and in
brackets of their own, (ThreeA:ThreeB!A1), ThreeA is a name, #NAME? unless
the workbook defines it.
"""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBARuntimeError

FIXTURES = Path(__file__).parent / 'fixtures'
PROBES: dict[str, dict[str, Any]] = {
    probe['name']: probe
    for fixture in ('three_d_references.json', 'three_d_functions.json')
    for probe in json.loads((FIXTURES / fixture).read_text(encoding='utf-8'))['probes']}

#: Probes the model does not answer as Excel does, and why.
GAPS: dict[str, str] = {
    'quoted_bracketed': "in brackets of their own Excel reads 'ThreeA:ThreeB'!A1 as a workbook named ThreeB, writing "
                        "it 'ThreeA:[ThreeB]ThreeB'!A1, #REF!; the model reads the 3D reference",
    'spaced_sheets_bracketed': "the same for 'Three A:Three B'!A1, which Excel writes 'Three A:[Three B]Three B'!A1",
}


@pytest.mark.parametrize('name', [
    pytest.param(name, marks=[pytest.mark.xfail(reason=GAPS[name], strict=True)] if name in GAPS else [])
    for name in PROBES])
def test_a_formula_given_a_3d_reference(name: str) -> None:
    probe = PROBES[name]
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(probe['macro'], name='Harness')
    if 'error' in probe:
        with pytest.raises(VBARuntimeError) as raised:
            app.run('Harness.Probe')
        assert raised.value.number == probe['error']
    else:
        assert app.run('Harness.Probe') == probe['result']


def test_every_gap_is_a_probe() -> None:
    assert set(GAPS) <= set(PROBES)
