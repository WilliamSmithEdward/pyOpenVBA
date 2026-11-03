"""Members of a range and a sheet, replayed from live Excel (scripts/measure_range_members.py,
tests/fixtures/range_members.json): what each answers or raises, and the cells it leaves."""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

RECORD: dict[str, Any] = json.loads(
    (Path(__file__).parent / 'fixtures/range_members.json').read_text(encoding='utf-8'))
NAMES = [case[0] for case in RECORD['cases']]
#: Left out: what the model has no object for, and what it refuses: several dependents on a sheet, whose order is not
#: understood, and Justify on runs of spaces. Justify with a line about as wide as its column is refused too, and only
#: where it is, since where Excel breaks those depends on the screen.
LEFT_OUT = ('.SoundNote', '.XPath', '.CellControl', '.ResetContents', 'd.Range("G3").Formula', 'runs  ')
KEPT = [index for index, name in enumerate(NAMES) if not any(text in name for text in LEFT_OUT)]


def _source() -> str:
    """The probe as recorded, with the statements of the cases left out taken away."""
    kept = set(KEPT)
    written: list[str] = []
    current = -1
    for line in RECORD['source'].split('\n'):
        stripped = line.strip()
        if line.startswith('    Case '):
            current = int(stripped[5:])
        elif stripped == 'End Select':
            current = -1
        elif current >= 0 and current not in kept:
            continue
        written.append(line)
    return '\n'.join(written)


@pytest.fixture(scope='module')
def run() -> dict[int, str]:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(_source(), name='MemberProbe')
    answers: dict[int, str] = {}
    for index in KEPT:
        try:
            answers[index] = str(app.run('MemberProbe.Probe', index, index))
        except VBAUnsupportedError as refused:
            assert 'about as wide as its column' in str(refused), NAMES[index]
    return answers


@pytest.mark.parametrize('index', KEPT, ids=[NAMES[index] for index in KEPT])
def test_each_case_as_excel_answered(run: dict[int, str], index: int) -> None:
    if index in run:
        assert run[index] == RECORD['cases'][index][1]


def test_several_dependents_on_a_sheet_are_refused() -> None:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module('''Public Sub Try()
    With ActiveWorkbook.Worksheets(1)
        .Range("B1").Formula = "=A1"
        .Range("C1").Formula = "=A1"
        .Range("A1").ShowDependents
        .Range("A1").NavigateArrow False, 1
    End With
End Sub
''', name='Trial')
    with pytest.raises(VBAUnsupportedError, match='several dependents'):
        app.run('Trial.Try')
