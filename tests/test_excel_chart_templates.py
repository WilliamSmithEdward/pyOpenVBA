"""Chart.SaveChartTemplate, ApplyChartTemplate, SetBackgroundPicture and Paste on a new embedded chart, against what
Excel answers and saves (scripts/measure_chart_templates.py, tests/fixtures/chart_templates.json).

A background picture tiles the chart area, the PNG kept byte for byte. Saving a template writes a picture of the chart
the model cannot render, and applying one is not built: both stay refused, as does Paste of a copied chart's formats,
which Excel itself failed to take from the clipboard. The probe's files are under TEMP, so the sandbox is off; a new
chart's unique ids and axis ids are random at Excel's save, so they are masked."""

from __future__ import annotations

import base64
import json
import re
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

FIXTURES = Path(__file__).parent / 'fixtures'
RECORD: dict[str, Any] = json.loads((FIXTURES / 'chart_templates.json').read_text(encoding='utf-8'))
REFUSED = ('save template', 'apply template', 'apply missing', 'paste formats')
_UNIQUE = re.compile(r'(uniqueId val="\{[0-9A-F]{8})-[0-9A-F-]+')
_AXIS = re.compile(r'(<c:(?:axId|crossAx) val=")\d+')


def what(case: list[Any]) -> str:
    return RECORD['edits'][case[0]][0]


def comparable(parts: dict[str, str]) -> dict[str, str]:
    return {name: _AXIS.sub(r'\1', _UNIQUE.sub(r'\1', text)) for name, text in parts.items()}


def run(case: list[Any], folder: Path) -> tuple[str, dict[str, str]]:
    picture, template = folder / 'background.png', folder / 'template.crtx'
    picture.write_bytes(base64.b64decode(RECORD['picture']))
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(RECORD['source'], name='P')
    out = folder / 'out.xlsx'
    answer = str(app.run('P.Act', str(out), case[0], str(template), str(picture)))
    with zipfile.ZipFile(out) as package:
        parts = {name: package.read(name) for name in package.namelist() if name.startswith(('xl/charts/', 'xl/media/'))}
    return answer, {name: data.decode('utf-8') if name.endswith(('.xml', '.rels')) else 'base64:' + base64.b64encode(data).decode()
                    for name, data in parts.items()}


@pytest.mark.parametrize('case', [case for case in RECORD['cases'] if what(case) in REFUSED], ids=what)
def test_an_edit_not_built_stays_refused(case: list[Any], tmp_path: Path) -> None:
    with pytest.raises(VBAUnsupportedError):
        run(case, tmp_path)


@pytest.mark.parametrize('case', [case for case in RECORD['cases'] if what(case) not in REFUSED], ids=what)
def test_what_is_answered_is_excels_answer(case: list[Any], tmp_path: Path) -> None:
    assert run(case, tmp_path)[0] == case[1]


@pytest.mark.parametrize('case', [case for case in RECORD['cases'] if what(case) not in REFUSED], ids=what)
def test_the_charts_save_as_excel_saved_them(case: list[Any], tmp_path: Path) -> None:
    assert comparable(run(case, tmp_path)[1]) == comparable(case[2])
