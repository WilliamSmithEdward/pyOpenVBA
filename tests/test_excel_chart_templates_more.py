"""Chart.SaveChartTemplate from four kinds of chart, and Chart.ApplyChartTemplate of each template onto three kinds of chart,
against what Excel writes and saves (scripts/measure_chart_templates_more.py, tests/fixtures/chart_templates_more.json).

A template's picture comes from the Export hook, so a stand-in function draws it here and the picture is not compared; every
other part is. A template is applied from the package Excel wrote, rebuilt from its parts. A new chart's unique ids and axis ids
are random at Excel's save, so they are masked."""
from __future__ import annotations

import base64
import json
import re
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.apps.excel._hooks import HookRequest
from pyopenvba.exceptions import VBAUnsupportedError

FIXTURES = Path(__file__).parent / 'fixtures'
RECORD: dict[str, Any] = json.loads((FIXTURES / 'chart_templates_more.json').read_text(encoding='utf-8'))
_UNIQUE = re.compile(r'(uniqueId val="\{[0-9A-F]{8})-[0-9A-F-]+')
_AXIS = re.compile(r'(<c:(?:axId|crossAx) val=")\d+')


def masked(xml: str) -> str:
    return _AXIS.sub(r'\1', _UNIQUE.sub(r'\1', xml))


def draw(request: HookRequest) -> None:
    assert request.path is not None
    request.path.write_bytes(b'BM stand-in')


def app_with(source: str) -> ExcelApplication:
    app = ExcelApplication(sandbox=False)
    app.set_hook('Export', draw)
    app.add_workbook()
    app.add_module(source, name='P')
    return app


def rebuilt(parts: dict[str, str], path: Path) -> None:
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as package:
        for name, text in parts.items():
            package.writestr(name, base64.b64decode(text[len('base64:'):]) if text.startswith('base64:') else text.encode('utf-8'))


@pytest.mark.parametrize('case', RECORD['templates'], ids=lambda case: RECORD['sources'][case[0]][0])
def test_a_template_holds_what_excel_writes(case: list[Any], tmp_path: Path) -> None:
    path = tmp_path / 't.crtx'
    assert str(app_with(RECORD['source']).run('P.MakeTemplate', case[0], str(path))) == case[1]
    with zipfile.ZipFile(path) as package:
        names = package.namelist()
        got = {name: package.read(name).decode('utf-8') for name in names if not name.endswith('.bmp')}
    assert names == list(case[2])
    assert {name: masked(text) for name, text in got.items()} == {
        name: masked(text) for name, text in case[2].items() if not name.endswith('.bmp')}


def test_a_template_needs_a_function_to_draw_it(tmp_path: Path) -> None:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(RECORD['source'], name='P')
    with pytest.raises(VBAUnsupportedError, match='set_hook'):
        app.run('P.MakeTemplate', 0, str(tmp_path / 't.crtx'))
    assert not (tmp_path / 't.crtx').exists()


def _label(case: list[Any]) -> str:
    return f"{RECORD['sources'][case[0]][0]} onto {RECORD['targets'][case[1]][0]}"


@pytest.mark.parametrize('case', RECORD['applied'], ids=_label)
def test_an_applied_template_saves_as_excel_saved_it(case: list[Any], tmp_path: Path) -> None:
    template, out = tmp_path / 't.crtx', tmp_path / 'out.xlsx'
    rebuilt(RECORD['templates'][case[0]][2], template)
    assert str(app_with(RECORD['source']).run('P.Apply', str(out), case[1], str(template))) == case[2]
    with zipfile.ZipFile(out) as package:
        got = {name: package.read(name).decode('utf-8') for name in package.namelist() if name.startswith(('xl/charts/', 'xl/theme/'))}
        types = package.read('[Content_Types].xml').decode('utf-8')
    want = {name: text for name, text in case[3].items() if name.startswith(('xl/charts/', 'xl/theme/'))}
    assert {name: masked(text) for name, text in got.items()} == {name: masked(text) for name, text in want.items()}
    # The model orders a package's content types otherwise than Excel for every workbook; the override's entry is Excel's.
    override = ('<Override PartName="/xl/theme/themeOverride1.xml" '
                'ContentType="application/vnd.openxmlformats-officedocument.themeOverride+xml"/>')
    assert override in types and override in case[3]['[Content_Types].xml']
