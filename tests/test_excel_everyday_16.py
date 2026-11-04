"""A sixteenth sweep of everyday macros, each run bound late and typed, replayed against live Excel
(scripts/measure_excel_everyday_16.py)."""
from __future__ import annotations

import json
import sys
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

PROBES: list[dict[str, Any]] = json.loads((Path(__file__).parent / 'fixtures/excel_everyday_16.json').read_text(encoding='utf-8'))['probes']
SEPARATOR = '\u00a6'
#: Lines the model answers otherwise for a stated reason. Range.Value of separate areas is the verification lane's, and a late read's
#: Double is the binding work's next step.
GAPS = ('TypeName(ws.Range("A1,C3").Value)', 'o.FileFormat', 'ColH()')


#: The macro joins TEMP and a name with a backslash, which is not how another system spells a path it answers.
SPELLED_BY_WINDOWS = ('o.FullName = f', 'o.Path = Environ')


def width_dependent(asked: str, answer: str) -> bool:
    """What .Text fills or hashes to depends on the column's pixels, and this was measured at 144 DPI where the model is 96."""
    return 'by ws.Ran' in asked and ('###' in answer or '* ' in asked)


@pytest.mark.parametrize('binding', ['late', 'typed'])
@pytest.mark.parametrize('probe', PROBES, ids=[probe['name'] for probe in PROBES])
def test_an_everyday_macro(probe: dict[str, Any], binding: str, monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    gaps = GAPS
    if sys.platform != 'win32':
        # The macro saves under TEMP, which only Windows sets.
        monkeypatch.setenv('TEMP', str(tmp_path))
        gaps = (*GAPS, *SPELLED_BY_WINDOWS)
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    # Each run has its own temporary files, as tests run side by side.
    app.add_module(probe[f'{binding}_macro'].replace('sweep16', f'sweep16{binding}').replace('SWEEP16', f'SWEEP16{binding.upper()}'), name='H')
    found = str(app.run('H.Probe')).replace(f'sweep16{binding}', 'sweep16').split(SEPARATOR)
    wanted = str(probe[binding]).split(SEPARATOR)
    assert len(found) == len(wanted)
    wrong = [(asked, excel, model) for asked, excel, model in zip(probe['asked'], wanted, found)
             if excel != model and not width_dependent(asked, excel) and not any(gap in asked for gap in gaps)]
    assert not wrong
