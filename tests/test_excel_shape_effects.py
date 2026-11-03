"""Shape.Shadow, Glow, SoftEdge and Reflection, replayed from live Excel (scripts/measure_shape_effects.py,
tests/fixtures/shape_effects.json): what each case reads, and the spPr its saved shape has."""
from __future__ import annotations

import json
import os
import re
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/shape_effects.json').read_text(encoding='utf-8'))
NAMES = [case[0] for case in RECORD['cases']]
#: A preset shadow, an inner one, and one set to rotate with its shape are refused.
REFUSED = ('sh.Shadow.Type', 'msoShadowStyleInnerShadow')
KEPT = [index for index, name in enumerate(NAMES) if not any(text in name for text in REFUSED)]


def drawn(path: Path) -> str:
    with zipfile.ZipFile(path) as package:
        xml = package.read('xl/drawings/drawing1.xml').decode('utf-8')
    found = re.search(r'<xdr:spPr\b.*?</xdr:spPr>(?:<xdr:style>.*?</xdr:style>)?', xml, re.DOTALL)
    return found.group(0) if found else ''


def _source() -> str:
    """The probe as recorded, with the statements of the refused cases taken away."""
    kept = set(KEPT)
    written: list[str] = []
    current = -1
    for line in RECORD['source'].replace('folder & "\\', f'folder & "{os.sep}').split('\n'):
        stripped = line.strip()
        if line.startswith('    Case '):
            current = int(stripped[5:])
        elif stripped == 'End Select':
            current = -1
        elif current >= 0 and current not in kept and not stripped.startswith('Set sh ='):
            continue
        written.append(line)
    return '\n'.join(written)


@pytest.fixture(scope='module')
def run(tmp_path_factory: pytest.TempPathFactory) -> tuple[dict[int, str], Path]:
    folder = tmp_path_factory.mktemp('shape_effects')
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(_source(), name='EffectProbe')
    return {index: str(app.run('EffectProbe.Probe', str(folder), index, index)) for index in KEPT}, folder


@pytest.mark.parametrize('index', KEPT, ids=[NAMES[index] for index in KEPT])
def test_each_case_as_excel_answered(run: tuple[dict[int, str], Path], index: int) -> None:
    assert run[0][index] == RECORD['cases'][index][1]


@pytest.mark.parametrize('index', KEPT, ids=[NAMES[index] for index in KEPT])
def test_each_shape_as_excel_saved_it(run: tuple[dict[int, str], Path], index: int) -> None:
    assert drawn(run[1] / f'c{index}.xlsx') == RECORD['cases'][index][2]


@pytest.mark.parametrize('statement', ['sh.Shadow.Type = msoShadow21', 'sh.Shadow.Style = msoShadowStyleInnerShadow',
                                       'sh.Shadow.RotateWithShape = msoTrue'])
def test_what_is_not_modelled_is_refused(statement: str) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('Public Sub Go()\n    Dim sh As Object\n'
                   '    Set sh = ActiveSheet.Shapes.AddShape(msoShapeRectangle, 10, 10, 100, 40)\n'
                   f'    {statement}\nEnd Sub\n', name='Refused')
    with pytest.raises(VBAUnsupportedError):
        app.run('Refused.Go')
