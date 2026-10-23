"""Which number formats make a Text cell, measured by scripts/measure_text_cells.py.

A Text cell keeps a formula, or a string that would read as a number or a
date, as the text written; a number written stays a number. A format
makes one where it has an @ and no section that shows digits or a date:
@, "x"@, General;;;@ and *-General;;;@ do, and 0;@, General;0;0;@,
m/d/yyyy;;;@ and the Comma style do not.
"""
from __future__ import annotations

import json
import re
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/text_cells.json').read_text(encoding='utf-8'))
PROBES: dict[str, dict[str, Any]] = {probe['name']: probe for probe in RECORD['probes']}
_REFUSED = 'Excel refuses an @ in two sections, after a condition or beside a date as error 1004; the model takes it'
_REWRITTEN = 'Excel rewrites ;@ as ;;;@ and 0;0;@ as 0;0;0;@ but keeps 0;@ and General;General;@; the model keeps all'
#: Formats whose NumberFormat Excel spells or refuses otherwise than the model, by why.
GAPS = {'@;@': _REFUSED, '[>5]@': _REFUSED, 'mm/dd @': _REFUSED, '@*-;@': _REFUSED,
        ';@': _REWRITTEN, ';@" kg"': _REWRITTEN, '0;0;@': _REWRITTEN}


def _formats(macro: str) -> list[str]:
    """The formats a probe sets, as its macro spells them."""
    return [text.replace('""', '"') for text in re.findall(r'^formats\(\d+\) = "(.*)"$', macro, re.MULTILINE)]


def _writes(macro: str) -> list[str]:
    return re.findall(r'^ +Case \d+: (.*)$', macro, re.MULTILINE)


@pytest.fixture(scope='module')
def replayed() -> dict[str, list[str]]:
    out: dict[str, list[str]] = {}
    for name, probe in PROBES.items():
        app = ExcelApplication()
        app.add_workbook()
        app.add_module(probe['macro'], name='Harness')
        out[name] = str(app.run('Harness.Probe')).split('|')
    return out


CASES = [pytest.param(name, index, id=f'{name}-{code}-{write}',
                       marks=[pytest.mark.xfail(reason=GAPS[code], strict=True)] if code in GAPS else [])
         for name, probe in PROBES.items()
         for index, (code, write) in enumerate((code, write) for code in _formats(probe['macro'])
                                               for write in _writes(probe['macro']))]


@pytest.mark.parametrize(('name', 'index'), CASES)
def test_a_write_to_a_formatted_cell(replayed: dict[str, list[str]], name: str, index: int) -> None:
    assert replayed[name][index] == PROBES[name]['result'].split('|')[index]
