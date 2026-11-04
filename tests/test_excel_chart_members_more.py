"""More chart members, replayed from live Excel (scripts/measure_chart_members_more.py,
tests/fixtures/chart_members_more.json): what each case reads, and the chart part it saves."""
from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError
from test_excel_chart_markup import chart_part

RECORD: dict[str, Any] = json.loads(
    (Path(__file__).parent / 'fixtures/chart_members_more.json').read_text(encoding='utf-8'))
NAMES = [case[0] for case in RECORD['cases']]
#: Refused, which each case below checks: a point's own shadow, the fonts of some of a title's characters, a filtered category, changing or selecting leader lines, a data table's Border and Format, and
#: where a slice lies, which needs the chart's layout.
REFUSED = ('.Points(2).Shadow =', '.Font.Bold', 'IsFiltered = True', 'LeaderLines.Border.Color = 255', 'LeaderLines.Select', 'DataTable.Border',
           'PieSliceLocation(xlHorizontalCoordinate) ', 'PieSliceLocation(xlVerticalCoordinate)')

def _refused(name: str) -> bool:
    if 'DataTable.Font.Bold' in name or 'LE.Font.Bold' in name or 'LegendEntries(1).Font.Bold' in name:
        return False
    if name.startswith('pie:') and 'PieSliceLocation(xlHorizontalCoordinate)' in name and 'xlInsideEnd' not in name:
        return True
    return any(text in name for text in REFUSED)


REFUSALS = [index for index, name in enumerate(NAMES) if _refused(name)]
KEPT = [index for index in range(len(NAMES)) if index not in REFUSALS]


def _source(kept: set[int]) -> str:
    """The probe as recorded, with the statements of the cases left out taken away."""
    written: list[str] = []
    current = -1
    for line in RECORD['source'].replace('folder & "\\', f'folder & "{os.sep}').split('\n'):
        stripped = line.strip()
        if line.startswith('    Case ') and not stripped.endswith(' which, ws, ch, r'):
            current = int(stripped[5:])
        elif stripped == 'End Select':
            current = -1
        elif current >= 0 and current not in kept:
            continue
        written.append(line)
    return '\n'.join(written)


@pytest.fixture(scope='module')
def run(tmp_path_factory: pytest.TempPathFactory) -> tuple[dict[int, str], Path]:
    folder = tmp_path_factory.mktemp('chart_members_more')
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(_source(set(KEPT)), name='MoreProbe')
    return {index: str(app.run('MoreProbe.Probe', str(folder), index, index)) for index in KEPT}, folder


@pytest.mark.parametrize('index', KEPT, ids=[NAMES[index] for index in KEPT])
def test_each_case_as_excel_answered(run: tuple[dict[int, str], Path], index: int) -> None:
    assert run[0][index] == RECORD['cases'][index][1]


@pytest.mark.parametrize('index', KEPT, ids=[NAMES[index] for index in KEPT])
def test_each_case_saves_the_part_excel_saved(run: tuple[dict[int, str], Path], index: int) -> None:
    assert chart_part(run[1] / f'c{index}.xlsx') == RECORD['cases'][index][2]


@pytest.mark.parametrize('index', REFUSALS, ids=[NAMES[index] for index in REFUSALS])
def test_what_is_not_built_says_so(tmp_path: Path, index: int) -> None:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(_source({index}), name='MoreProbe')
    with pytest.raises(VBAUnsupportedError):
        app.run('MoreProbe.Probe', str(tmp_path), index, index)
