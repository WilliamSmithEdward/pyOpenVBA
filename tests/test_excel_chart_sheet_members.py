"""Chart.PrintedCommentPages and Chart.SaveAs on a chart sheet and an embedded chart, and a chart sheet's Activate and
Deactivate events with the workbook's, as native Excel answers and raises them (scripts/measure_chart_sheet_members.py,
tests/fixtures/chart_sheet_members.json). The probes save under TEMP, a folder of the test's own, so the sandbox is off."""
from __future__ import annotations

import json
from pathlib import Path

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBARuntimeError

FIXTURE = json.loads((Path(__file__).parent / 'fixtures/chart_sheet_members.json').read_text(encoding='utf-8'))
PROBES: list[dict[str, object]] = FIXTURE['probes']
CLASSES: dict[str, str] = FIXTURE['classes']


@pytest.mark.parametrize('probe', PROBES, ids=[str(probe['name']) for probe in PROBES])
def test_as_native_excel(probe: dict[str, object], tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    # The probes save under Environ("TEMP"), which is set on Windows alone: here it is a folder of the test's own.
    monkeypatch.setenv('TEMP', str(tmp_path))
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    for name, source in CLASSES.items():
        app.add_module(source, name=name, kind='class')
    app.add_module(str(probe['macro']), name='Harness')
    if 'error' in probe:
        with pytest.raises(VBARuntimeError) as raised:
            app.run('Harness.Probe')
        assert raised.value.number == probe['error']
    else:
        assert app.run('Harness.Probe') == probe['result']
