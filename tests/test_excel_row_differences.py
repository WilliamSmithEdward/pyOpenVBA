"""Range.RowDifferences and ColumnDifferences, over the grid and the ranges live Excel was asked.

Measured by scripts/measure_row_differences.py: each cell's FormulaR1C1 is
compared with the comparison cell's, without regard to case; the comparison
is its first area's last cell, one row high, inside the range.
"""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBARuntimeError

RECORD: dict[str, Any] = json.loads(
    (Path(__file__).parent / 'fixtures/row_differences.json').read_text(encoding='utf-8'))


@pytest.fixture(scope='module')
def app() -> ExcelApplication:
    made = ExcelApplication()
    made.add_workbook()
    lines = ['Sub Lay()', 'Dim ws As Worksheet', 'Set ws = ActiveSheet']
    lines += [f'ws.Range("{cell}").Formula = "' + value.replace('"', '""') + '"' for cell, value in RECORD['grid'].items()]
    lines += RECORD['after'] + ['End Sub']
    made.add_module('\n'.join(lines) + '\n', name='Grid')
    made.run('Grid.Lay')
    return made


def test_what_each_rival_reads(app: ExcelApplication) -> None:
    """FormulaR1C1, Text and Value of the cells that split the rival readings, as Excel read them."""
    app.add_module('Function Read() As String\nDim c As Range, s As String\nFor Each c In ActiveSheet.Range("A10:D14")\n'
                   's = s & "|" & c.Address(False, False) & "=" & c.FormulaR1C1 & "~" & c.Text & "~" & CStr(c.Value)\n'
                   'Next\nRead = Mid(s, 2)\nEnd Function\n', name='Readings')
    assert str(app.run('Readings.Read')).split('|') == RECORD['readings']


@pytest.mark.parametrize('case', list(RECORD['cases']))
def test_differences_as_excel_finds_them(app: ExcelApplication, case: str) -> None:
    app.add_module(f'Function Found() As String\nDim r As Range\nSet r = ActiveSheet.{case}\n'
                   'Found = r.Address(False, False) & "~" & r.Areas.Count\nEnd Function\n', name='Asking')
    want = RECORD['cases'][case]
    if want.startswith('Err '):
        number, _, description = want[len('Err '):].partition(' ')
        with pytest.raises(VBARuntimeError) as raised:
            app.run('Asking.Found')
        assert raised.value.number == int(number)
        assert description in str(raised.value)
    else:
        assert app.run('Asking.Found') == want
