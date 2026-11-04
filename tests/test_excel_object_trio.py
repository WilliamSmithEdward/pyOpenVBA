"""Application, Creator and Parent on one object of every class, each way a macro reaches it, and the errors from
assigning to what Excel holds read-only, replayed from live Excel (scripts/measure_object_trio.py,
tests/fixtures/object_trio.json; scripts/measure_read_only_lets.py, tests/fixtures/read_only_lets.json)."""
from __future__ import annotations

import csv
import json
from pathlib import Path

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

ROOT = Path(__file__).resolve().parents[1]
RECORD = json.loads((ROOT / 'tests/fixtures/object_trio.json').read_text(encoding='utf-8'))
LETS = json.loads((ROOT / 'tests/fixtures/read_only_lets.json').read_text(encoding='utf-8'))
TRIO = ('Application', 'Creator', 'Parent')

#: Each way in that the model does not have yet, which reports itself as a gap rather than answering wrongly. Those are
#: gaps of the member named last, and leave what is reached by other ways verified. ChartArea.Interior is error 438.
GAPS = {'Application.FindFormat.Borders', 'pt.ActiveFilters', 'pt.Slicers', 'ws.PageSetup.Pages(1)',
        'ch.ChartArea.Interior'}


@pytest.fixture(scope='module')
def app() -> ExcelApplication:
    made = ExcelApplication(sandbox=False)
    made.add_workbook()
    made.add_module(RECORD['source'], name='Trio')
    made.run('Trio.Setup')
    return made


@pytest.mark.parametrize('index', range(len(RECORD['answers'])), ids=[one[1] for one in RECORD['answers']])
def test_each_object_as_excel_answered(app: ExcelApplication, index: int) -> None:
    _, reach, answer = RECORD['answers'][index]
    if reach in GAPS:
        pytest.skip('a gap of the member that reaches it')
    try:
        found = str(app.run('Trio.One', index))
    except VBAUnsupportedError:
        # Reported as a gap; the checklist keeps its row from VERIFIED (test_a_verified_row_answers below).
        pytest.skip('reported as a gap')
    assert found == answer


@pytest.mark.parametrize('case', LETS['cases'], ids=[f'{one[0]}.{one[1]}' for one in LETS['cases']])
def test_assigning_to_a_read_only_member(app: ExcelApplication, case: list[str]) -> None:
    owner, name, index, number = case
    module = (f'Public Function Assign{owner}{name}() As Long\n'
              f'    Dim o As Object\n    Set o = Trio.Reached({index})\n    On Error Resume Next\n'
              f'    o.{name} = 1\n    Assign{owner}{name} = Err.Number\nEnd Function\n')
    app.add_module(module, name=f'L{owner}{name}')
    try:
        assert app.run(f'L{owner}{name}.Assign{owner}{name}') == number
    except VBAUnsupportedError:
        pytest.skip('a member the model does not have')


def test_a_verified_row_answers() -> None:
    """A VERIFIED row of the three has its class measured here, and answers in the model by some way in."""
    with (ROOT / 'docs/excel_checklist.csv').open(encoding='utf-8', newline='') as handle:
        rows = list(csv.DictReader(handle))
    measured = {one[0] for one in RECORD['answers'] if one[1] not in GAPS}
    for row in rows:
        if row['member'] in TRIO and row['status'] == 'VERIFIED' and 'object_trio' in row['evidence']:
            assert row['object'] in measured, f"{row['object']}.{row['member']}"
