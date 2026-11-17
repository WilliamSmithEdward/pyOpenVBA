"""ShapeRange's PictureFormat, Nodes, Vertices, Callout and _Default over one shape and over two; Application, Creator and
Parent on the shape effect formats and a header picture's Graphic; and a Graphic's looks on a sheet with no picture,
replayed from live Excel (scripts/measure_shape_format_rows.py, tests/fixtures/shape_format_rows.json).

Each read is replayed on its own, from the statements the probe recorded, so that one the model refuses does not hide the
others. What the model refuses is named here and nothing else may be."""
from __future__ import annotations

import json
import re
from pathlib import Path

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

FIXTURES = Path(__file__).parent / 'fixtures'
RECORD = json.loads((FIXTURES / 'shape_format_rows.json').read_text(encoding='utf-8'))
NAMES = [name for name, _ in RECORD['ranges']]
#: Each case's statements, as the probe ran them.
STATEMENTS = re.findall(r'^    Case \d+\n        (.+)$', RECORD['source'].split('Private Function Graphics')[0], re.MULTILINE)
#: A new AutoShape's nodes are its preset's, which the model does not draw; TextEffect and ThreeD are not modelled on Shape.
REFUSED = ({(name, read) for name in ('rect', 'callout', 'two') for read in ('r.Nodes.Count', 'sh.Nodes.Count')}
           - {('two', 'r.Nodes.Count')}
           | {(name, read) for name in NAMES for read in RECORD['reads'] if 'TextEffect' in read or 'ThreeD' in read})


def png(folder: Path) -> None:
    import struct
    import zlib

    def chunk(tag: bytes, data: bytes) -> bytes:
        return struct.pack('>I', len(data)) + tag + data + struct.pack('>I', zlib.crc32(tag + data))
    rows = b''.join(b'\x00' + b'\xff\x00\x00' * 8 for _ in range(4))
    (folder / 'a.png').write_bytes(b''.join([b'\x89PNG\r\n\x1a\n', chunk(b'IHDR', struct.pack('>IIBBBBB', 8, 4, 8, 2, 0, 0, 0)),
                                             chunk(b'IDAT', zlib.compress(rows, 9)), chunk(b'IEND', b'')]))


def _one_read(which: int, read: str) -> str:
    return '\n'.join([
        'Public Function Read(ByVal folder As String) As String',
        '    Dim wb As Workbook, ws As Worksheet, r As ShapeRange, sh As Shape, o As Object, b As Object, v As String',
        '    Set wb = Workbooks.Add(xlWBATWorksheet): Set ws = wb.Worksheets(1)',
        f'    {STATEMENTS[which]}', '    Set sh = r(1): Set o = r', '    On Error Resume Next',
        f'    v = {read}', '    If Err.Number Then v = "E" & Err.Number', '    Read = v', 'End Function', ''])


CASES = [(which, index) for which in range(len(NAMES)) for index in range(len(RECORD['reads']))]


@pytest.mark.parametrize('case', CASES, ids=[f"{NAMES[which]}: {RECORD['reads'][index]}" for which, index in CASES])
def test_each_range_read_as_excel_answered(case: tuple[int, int], tmp_path: Path) -> None:
    which, index = case
    name, read = NAMES[which], RECORD['reads'][index]
    png(tmp_path)
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(_one_read(which, read), name='P')
    try:
        found = str(app.run('P.Read', str(tmp_path)))
    except VBAUnsupportedError:
        assert (name, read) in REFUSED, f'{name}: {read} is refused'
        return
    assert (name, read) not in REFUSED, f'{name}: {read} now answers; take it out of REFUSED'
    assert found == RECORD['ranges'][which][1][index]


def _public(source: str) -> str:
    return source.replace('Private Function Graphics', 'Public Function Graphics').replace(
        'Private Function Trios', 'Public Function Trios')


def test_a_graphic_with_no_picture() -> None:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(_public(RECORD['source']), name='P')
    assert str(app.run('P.Graphics')).split('|') == RECORD['graphic']['answers']


def test_the_effects_and_graphics_application_creator_and_parent() -> None:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(_public(RECORD['source']), name='P')
    found = str(app.run('P.Trios')).split('@')
    assert found == [answer for _, _, answer in RECORD['trio']]
