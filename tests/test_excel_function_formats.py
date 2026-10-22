"""The number format each function gives a General cell, measured by scripts/measure_function_formats.py.

Of the engine's 501 functions, Excel gives a cell a format for 18: DATE and
TODAY m/d/yyyy, NOW m/d/yyyy h:mm and TIME h:mm AM/PM; DB, DDB, FV, IPMT,
NPV, PMT, PPMT, PV, SLN, SYD and VDB the currency format; IRR, MIRR and RATE
0%. Inside another formula a function's format goes as a reference's does:
NPV(..)+1 is currency, NPV(..)*2 is General.
"""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

RECORD: dict[str, Any] = json.loads(
    (Path(__file__).parent / 'fixtures/function_formats.json').read_text(encoding='utf-8'))
#: Each formula Excel was given, and the NumberFormat it gave the cell, or !1004 where it refused the formula.
CASES: dict[str, str] = {
    **{one['formula']: one['format'] for one in RECORD['functions'].values()},
    **RECORD['formulas'],
}


def _formats(formulas: list[str]) -> list[str]:
    """What the model gives a General cell each formula is written to, as the measuring macro reads it."""
    lines = ['Public Function Probe() As String', 'Dim ws As Object, out As String',
             'Set ws = ActiveSheet', 'ws.Range("A1").Value = 1', 'On Error Resume Next']
    for formula in formulas:
        text = formula.replace('"', '""')
        lines += ['ws.Range("H1").Clear', 'Err.Clear', f'ws.Range("H1").Formula = "{text}"',
                  'If Err.Number <> 0 Then out = out & "!" & Err.Number & vbLf Else '
                  'out = out & ws.Range("H1").NumberFormat & vbLf']
    lines += ['On Error GoTo 0', 'Probe = out', 'End Function']
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('\n'.join(lines) + '\n', name='Harness')
    return str(app.run('Harness.Probe')).split('\n')[:-1]


@pytest.fixture(scope='module')
def found() -> dict[str, str]:
    formulas = list(CASES)
    return dict(zip(formulas, _formats(formulas)))


@pytest.mark.parametrize('formula', list(CASES))
def test_the_format_a_formula_brings(found: dict[str, str], formula: str) -> None:
    assert found[formula] == CASES[formula]


@pytest.mark.parametrize('formula', list(RECORD['readback']))
def test_the_cell_reads_back_as_excel_reads_it(formula: str) -> None:
    """A currency cell's Value is a Currency, to four places, and a 0% cell's a Double."""
    app = ExcelApplication()
    app.add_workbook()
    text = formula.replace('"', '""')
    app.add_module(f'Function Probe() As String\nRange("H1").Formula = "{text}"\n'
                   'Probe = TypeName(Range("H1").Value) & "|" & CStr(Range("H1").Value) & "|" & Range("H1").Text & "|" '
                   '& CStr(Range("H1").Value2)\nEnd Function\n', name='Harness')
    assert app.run('Harness.Probe') == RECORD['readback'][formula]
