"""A chart's protection and page setup, replayed from live Excel (scripts/measure_chart_protection.py,
tests/fixtures/chart_protection.json): what each case reads, and the chart part it saves."""
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
    (Path(__file__).parent / 'fixtures/chart_protection.json').read_text(encoding='utf-8'))
NAMES = [case[0] for case in RECORD['cases']]
#: Refused, which each case below checks: a bevel, the effects of data labels, a page's print quality, and edits to a chart on
#: a protected sheet.
EFFECTS = ('.Format.ThreeD.BevelTopType = ',)
LABELS = ('DataLabels.Format.Glow', 'DataLabels.Format.Shadow', 'DataLabels.Format.SoftEdge', 'DataLabels.Format.ThreeD', 'DataLabels.Format.TextFrame2', 'DataLabels.Format.Adjustments', 'DataLabels.Format.PictureFormat', 'DataLabels.Format.AutoShapeType')
PROTECTED_EDITS = ('ws.Protect: ch.SeriesCollection', 'ws.Protect: ch.HasTitle', 'ws.Protect: ch.ChartType', 'ws.Protect: ch.HasLegend', 'ws.Protect: ch.SetSourceData', 'ws.Protect: ch.ChartArea', 'ws.Protect: ch.Parent')
REFUSED = (*EFFECTS, *LABELS, *PROTECTED_EDITS, 'PageSetup.PrintQuality')
#: A pie's automatic title, which working out a read of its page brings about in Excel and not here.
UNTAKEN = 'pie: ch.PageSetup.Orientation = xlLandscape'
#: Effects a line's or axis's Format does not have, which answer Excel's error before any setting.
ANSWERED = ('.ErrorBars.Format.ThreeD', '.Trendlines(1).Format.ThreeD', 'ch.Axes(xlValue).Format.ThreeD', 'MajorGridlines.Format.ThreeD', 'TickLabels.Format.ThreeD')
REFUSALS = [index for index, name in enumerate(NAMES)
            if any(text in name for text in REFUSED) and not any(text in name for text in ANSWERED)]
KEPT = [index for index in range(len(NAMES)) if index not in REFUSALS and not NAMES[index].startswith(UNTAKEN)]


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
    folder = tmp_path_factory.mktemp('chart_protection')
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(_source(set(KEPT)), name='ProtectionProbe')
    return {index: str(app.run('ProtectionProbe.Probe', str(folder), index, index)) for index in KEPT}, folder


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
    app.add_module(_source({index}), name='ProtectionProbe')
    with pytest.raises(VBAUnsupportedError):
        app.run('ProtectionProbe.Probe', str(tmp_path), index, index)
