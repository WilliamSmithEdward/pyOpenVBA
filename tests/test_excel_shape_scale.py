"""Shape.ScaleWidth and ScaleHeight, ConnectionSiteCount and the style reads, and ShapeRange's answers for its shapes,
replayed from live Excel (scripts/measure_shape_scale.py, tests/fixtures/shape_scale.json)."""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/shape_scale.json').read_text(encoding='utf-8'))
#: Cases the model refuses: scaling a rotated shape, which Excel turns about a point the model does not work out, and
#: the text frame of several shapes at once. The replay runs them with no statements and leaves their answers out.
REFUSED = ('sh.Rotation = 90', 'two: TypeName(sr.TextFrame')
RANGES = [name.split(': ', 1)[1] for name, _ in RECORD['cases'] if name.startswith(('one: ', 'two: ', 'chart: '))]
FIRST_RANGE = next(index for index, (name, _) in enumerate(RECORD['cases']) if name.startswith('one: '))


def _source() -> str:
    source: str = RECORD['source']
    for index, (name, _) in enumerate(RECORD['cases']):
        if not name.startswith(REFUSED):
            continue
        if index >= FIRST_RANGE:
            place = index - FIRST_RANGE
            source = source.replace(f'    Case {place}\n        s = {RANGES[place]}\n', f'    Case {place}\n', 1)
        else:
            source = source.replace(f'        {name}\n', '', 1)
    return source


KEPT = [index for index, (name, _) in enumerate(RECORD['cases']) if not name.startswith(REFUSED)]


@pytest.fixture(scope='module')
def answers() -> list[str]:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(_source(), name='ScaleProbe')
    return str(app.run('ScaleProbe.Probe')).split('@@@')


@pytest.mark.parametrize('index', KEPT, ids=[RECORD['cases'][index][0] for index in KEPT])
def test_each_case_as_excel_answered(answers: list[str], index: int) -> None:
    assert answers[index] == RECORD['cases'][index][1]


def test_scaling_a_rotated_shape_is_refused() -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('Public Sub Go()\n    Dim sh As Object\n'
                   '    Set sh = ActiveSheet.Shapes.AddShape(1, 100, 50, 100, 40)\n'
                   '    sh.Rotation = 90\n    sh.ScaleWidth 2, msoFalse\nEnd Sub\n', name='Rotated')
    with pytest.raises(VBAUnsupportedError, match='rotated'):
        app.run('Rotated.Go')
