"""Members of a chart macros reach for, replayed from live Excel (scripts/measure_chart_members.py,
tests/fixtures/chart_members.json): what each answers or raises, and the sheets and charts it leaves."""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

RECORD: dict[str, Any] = json.loads(
    (Path(__file__).parent / 'fixtures/chart_members.json').read_text(encoding='utf-8'))
NAMES = [case[0] for case in RECORD['cases']]
#: Left out: chart sheets, which the model does not hold; moving a chart between sheets; a chart's own shapes; 3-D chart
#: types; the axis border and the series' ClearFormats, whose markup was not measured.
LEFT_OUT = ('wb.Charts', 'ch.Location(xlLocationAutomatic)', 'Location(xlLocationAsObject, "Other")',
            'Location(xlLocationAsNewSheet', 'Shapes.Add', 'xl3DColumnClustered', '.Border', 'MinorUnitScale',
            'ClearFormats', 'LeaderLines', 'ch.Parent.Delete', 'co.Delete')
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
    return {index: str(app.run('MemberProbe.Probe', index, index)) for index in KEPT}


@pytest.mark.parametrize('index', KEPT, ids=[NAMES[index] for index in KEPT])
def test_each_case_as_excel_answered(run: dict[int, str], index: int) -> None:
    assert run[index] == RECORD['cases'][index][1]
