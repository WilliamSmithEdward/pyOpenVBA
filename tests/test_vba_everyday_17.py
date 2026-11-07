"""A seventeenth sweep: the VBA language as everyday macros use it, replayed against live Excel (scripts/measure_vba_everyday_17.py)."""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

PROBES: list[dict[str, Any]] = json.loads((Path(__file__).parent / 'fixtures/vba_everyday_17.json').read_text(encoding='utf-8'))['probes']
SEPARATOR = '\u00a6'
#: An Enum after a module's procedures, which the editor would refuse and adding a module lets in: Excel reads its names as undeclared.
GAPS = ('F20()',)


@pytest.mark.parametrize('probe', [probe for probe in PROBES if probe['name'] != 'collections'], ids=lambda probe: probe['name'])
def test_a_group_of_lines_answers_as_excel_did(probe: dict[str, Any]) -> None:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(probe['macro'], name='H')
    found = str(app.run('H.Probe')).split(SEPARATOR)
    wanted = str(probe['answer']).split(SEPARATOR)
    assert len(found) == len(wanted)
    assert not [(asked, excel, model) for asked, excel, model in zip(probe['asked'], wanted, found) if excel != model and asked not in GAPS]


def test_collections_answer_as_excel_did_but_for_a_class_that_is_not_there() -> None:
    probe = next(one for one in PROBES if one['name'] == 'collections')
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(probe['macro'], name='H')
    # CreateObject of no class stops the model, as it cannot tell that from one it has no server for. Excel says 429.
    with pytest.raises(VBAUnsupportedError):
        app.run('H.C18')
    for at, asked in enumerate(probe['asked']):
        if asked == 'C18()':
            continue
        wanted = str(probe['answer']).split(SEPARATOR)[at]
        try:
            found = app.run(f'H.{asked[:-2]}')
        except Exception as failed:  # noqa: BLE001 - a line stopping is its error number, as the macro reads it under On Error
            found = failed
        assert wanted.split(':', 1)[-1] == str(found) or wanted.split('=', 1)[0] == str(getattr(found, 'number', '')), asked
