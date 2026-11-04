"""Shapes.AddPicture of a PNG kept in the workbook, replayed from live Excel (scripts/measure_shape_add_picture.py,
tests/fixtures/shape_add_picture.json)."""
from __future__ import annotations

import base64
import json
import os
import re
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/shape_add_picture.json').read_text(encoding='utf-8'))
#: A picture's fill and line report themselves for now, as does a linked picture; a crop reads through Crop, not built.
UNBUILT_READS = ('sh.Fill.', 'sh.Line.', 'sh.PictureFormat.Crop.')
LEFT_OUT = {index for index, case in enumerate(RECORD['cases'])
            if 'msoTrue, msoFalse' in case[0] or 'msoTrue, msoTrue' in case[0] or 'sh.Line.' in case[0] or 'sh.Fill.' in case[0]
            or 'sh.Delete' in case[0]}
KEPT = [index for index in range(len(RECORD['cases'])) if index not in LEFT_OUT]
READS = [index + 1 for index, read in enumerate(RECORD['reads']) if not read.startswith(UNBUILT_READS)]
STRIP = re.compile(r'<a:ext uri="\{FF2B5EF4[^>]*>.*?</a:ext>|<a:extLst></a:extLst>')


def _source() -> str:
    source = RECORD['source'].replace('folder & "\\', f'folder & "{os.sep}').replace('out & "\\', f'out & "{os.sep}')
    for read in RECORD['reads']:
        if read.startswith(UNBUILT_READS):
            source = source.replace(f'    Err.Clear\n    v = ""\n    v = {read}\n    If Err.Number Then v = "E" & Err.Number\n'
                                    '    s = s & "," & v\n', '')
    for index in LEFT_OUT:
        source = re.sub(rf'    Case {index}\n        [^\n]*\n', f'    Case {index}\n        Set sh = Nothing\n', source)
    return source


def kept_reads(answer: str) -> str:
    parts = answer.split(',')
    return ','.join([parts[0], *(parts[index] for index in READS)])


@pytest.fixture(scope='module')
def run(tmp_path_factory: pytest.TempPathFactory) -> tuple[list[str], Path, Path]:
    images, out = tmp_path_factory.mktemp('images'), tmp_path_factory.mktemp('out')
    for name, data in RECORD['images'].items():
        (images / name).write_bytes(base64.b64decode(data))
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(_source(), name='PictureProbe')
    return str(app.run('PictureProbe.Probe', str(images), str(out))).split('@@@'), images, out


@pytest.mark.parametrize('index', KEPT, ids=[RECORD['cases'][index][0][-60:] for index in KEPT])
def test_each_case_as_excel_answered(run: tuple[list[str], Path, Path], index: int) -> None:
    assert run[0][index] == kept_reads(RECORD['cases'][index][1])


@pytest.mark.parametrize('index', KEPT, ids=[RECORD['cases'][index][0][-60:] for index in KEPT])
def test_each_workbook_as_excel_saved_it(run: tuple[list[str], Path, Path], index: int) -> None:
    case = RECORD['cases'][index][2]
    with zipfile.ZipFile(run[2] / f'c{index}.xlsx') as package:
        media = {name: base64.b64encode(package.read(name)).decode('ascii') for name in package.namelist() if '/media/' in name}
        assert media == case['media']
        if 'xl/drawings/drawing1.xml' not in case:
            return
        drawing = package.read('xl/drawings/drawing1.xml').decode('utf-8')
        assert STRIP.sub('', drawing) == STRIP.sub('', case['xl/drawings/drawing1.xml'])
        assert package.read('xl/drawings/_rels/drawing1.xml.rels').decode('utf-8') == case['xl/drawings/_rels/drawing1.xml.rels']
        assert '<Default Extension="png" ContentType="image/png"/>' in package.read('[Content_Types].xml').decode('utf-8')


def test_a_sandboxed_macro_cannot_read_the_file(tmp_path: Path) -> None:
    from pyopenvba.exceptions import VBASandboxError

    app = ExcelApplication()
    app.add_workbook()
    app.add_module(f'Public Sub Go()\n    ActiveSheet.Shapes.AddPicture "{tmp_path / "a.png"}", 0, -1, 1, 1, 9, 9\nEnd Sub\n', name='Sand')
    with pytest.raises(VBASandboxError):
        app.run('Sand.Go')


def test_other_images_and_links_report_themselves(tmp_path: Path) -> None:
    (tmp_path / 'a.txt').write_bytes(b'not an image')
    (tmp_path / 'b.png').write_bytes(base64.b64decode(RECORD['images']['a.png']))
    for statement in (f'"{tmp_path / "a.txt"}", 0, -1', f'"{tmp_path / "b.png"}", -1, 0'):
        app = ExcelApplication(sandbox=False)
        app.add_workbook()
        app.add_module(f'Public Sub Go()\n    ActiveSheet.Shapes.AddPicture {statement}, 1, 1, 9, 9\nEnd Sub\n', name='Other')
        with pytest.raises(VBAUnsupportedError):
            app.run('Other.Go')
