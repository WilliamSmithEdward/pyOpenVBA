"""Shapes.AddPicture of JPEG and GIF files, and PictureFormat, replayed from live Excel
(scripts/measure_shape_picture_formats.py, tests/fixtures/shape_picture_formats.json)."""
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

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/shape_picture_formats.json').read_text(encoding='utf-8'))
#: Excel converts a BMP, embeds what is no image as a stand-in, and moves a shape by rules not worked out on a crop.
LEFT_OUT = ('.bmp', 'text.png', 'PictureFormat.Crop')
CASES = [case for case in RECORD['cases'] if not any(one in case[0] for one in LEFT_OUT)]
IDS = [f'{index}: {re.sub(r"ws.Shapes.AddPicture.folder & ", "", case[0])}' for index, case in enumerate(CASES)]
#: The cells an anchor lands in follow the row height of the display it was measured on (14.5 points here, not the model's 15).
STRIP = re.compile(r'<a:ext uri="\{FF2B5EF4[^>]*>.*?</a:ext>|<a:extLst></a:extLst>|<xdr:from>.*?</xdr:to>')


@pytest.fixture(scope='module')
def images(tmp_path_factory: pytest.TempPathFactory) -> Path:
    folder = tmp_path_factory.mktemp('images')
    for name, data in RECORD['images'].items():
        (folder / name).write_bytes(base64.b64decode(data))
    return folder


@pytest.mark.parametrize('case', CASES, ids=IDS)
def test_each_case_as_excel_answered(case: list[Any], images: Path, tmp_path: Path) -> None:
    lines = ['Public Function Go(ByVal folder As String) As String', '    Dim ws As Object, sh As Object, s As String, v As String', '    Set ws = ActiveSheet',
             '    On Error GoTo Failed', f'    {case[0]}'.replace('\\', __import__('os').sep), '    s = "ok"', 'Measured:', '    On Error Resume Next']
    for read in RECORD['reads']:
        lines += ['    Err.Clear', '    v = ""', f'    v = {read}', '    If Err.Number Then v = "E" & Err.Number', '    s = s & "|" & v']
    lines += ['    Go = s', '    Exit Function', 'Failed:', '    s = "Err " & Err.Number', '    Resume Measured', 'End Function']
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module('\n'.join(lines) + '\n', name='Pictures')
    assert app.run('Pictures.Go', str(images)) == case[1]
    path = tmp_path / 'b.xlsx'
    app.save(path)
    want = case[2]
    with zipfile.ZipFile(path) as package:
        media = {name: base64.b64encode(package.read(name)).decode('ascii') for name in package.namelist() if '/media/' in name}
        assert media == want['media']
        if not media:
            return
        drawing = package.read('xl/drawings/drawing1.xml').decode('utf-8')
        assert STRIP.sub('', drawing) == STRIP.sub('', want['xl/drawings/drawing1.xml'])
        assert package.read('xl/drawings/_rels/drawing1.xml.rels').decode('utf-8') == want['xl/drawings/_rels/drawing1.xml.rels']
        assert re.findall(r'<Default [^>]*>', package.read('[Content_Types].xml').decode('utf-8')) == re.findall(r'<Default [^>]*>', want['[Content_Types].xml'])


@pytest.mark.parametrize('statements', [
    'sh.PictureFormat.CropLeft = 10', 'sh.PictureFormat.Crop.ShapeWidth = 50',
])
def test_what_was_not_worked_out_reports_itself(statements: str, images: Path) -> None:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(f'Public Sub Go()\n    Dim sh As Object\n    Set sh = ActiveSheet.Shapes.AddPicture("{images / "shown.png"}", 0, -1, 1, 1, 9, 9)\n'
                   f'    {statements}\nEnd Sub\n', name='Unbuilt')
    with pytest.raises(VBAUnsupportedError):
        app.run('Unbuilt.Go')


@pytest.mark.parametrize('name', ['zero.bmp', 'text.png'])
def test_a_bmp_and_what_is_no_image_report_themselves(name: str, images: Path) -> None:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(f'Public Sub Go()\n    ActiveSheet.Shapes.AddPicture "{images / name}", 0, -1, 1, 1, 9, 9\nEnd Sub\n', name='Other')
    with pytest.raises(VBAUnsupportedError):
        app.run('Other.Go')
