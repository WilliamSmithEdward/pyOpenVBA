"""Worksheet.SetBackgroundPicture replayed from live Excel (scripts/measure_background_picture.py,
tests/fixtures/background_picture.json): what it answers, and the picture element, relationships, media parts and
default content types the saved workbook holds."""
from __future__ import annotations

import json
import os
import re
import struct
import zipfile
import zlib
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

RECORD: dict[str, Any] = json.loads(
    (Path(__file__).parent / 'fixtures/background_picture.json').read_text(encoding='utf-8'))
#: The probe as recorded, its files joined to the folder with this platform's separator.
SOURCE = RECORD['source'].replace('folder & "\\', f'folder & "{os.sep}')


def red_png() -> bytes:
    def chunk(kind: bytes, data: bytes) -> bytes:
        return struct.pack('>I', len(data)) + kind + data + struct.pack('>I', zlib.crc32(kind + data) & 0xFFFFFFFF)
    raw = b''.join(b'\x00' + b'\xff\x00\x00' * 2 for _ in range(2))
    return (b'\x89PNG\r\n\x1a\n' + chunk(b'IHDR', struct.pack('>IIBBBBB', 2, 2, 8, 2, 0, 0, 0))
            + chunk(b'IDAT', zlib.compress(raw, 9)) + chunk(b'IEND', b''))


def saved(path: Path, picture: bytes) -> dict[str, object]:
    with zipfile.ZipFile(path) as package:
        names = package.namelist()
        sheets = sorted(name for name in names if re.fullmatch(r'xl/worksheets/sheet\d+\.xml', name))
        found: dict[str, object] = {}
        for sheet in sheets:
            text = package.read(sheet).decode('utf-8')
            rels = sheet.replace('worksheets/', 'worksheets/_rels/') + '.rels'
            found[sheet] = [text[text.find('<pageMargins'):],
                            package.read(rels).decode('utf-8') if rels in names else '']
        media = sorted(name for name in names if name.startswith('xl/media/'))
        found['media'] = [[name, package.read(name) == picture] for name in media]
        found['defaults'] = re.findall(r'<Default\b[^>]*/>', package.read('[Content_Types].xml').decode('utf-8'))
    return found


@pytest.fixture(scope='module')
def run(tmp_path_factory: pytest.TempPathFactory) -> tuple[list[str], Path]:
    folder = tmp_path_factory.mktemp('background_picture')
    (folder / 'bg.png').write_bytes(red_png())
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(SOURCE, name='BackgroundProbe')
    return str(app.run('BackgroundProbe.Probe', str(folder))).split('@@@'), folder


@pytest.mark.parametrize('index', range(len(RECORD['cases'])))
def test_each_case_as_excel_answered(run: tuple[list[str], Path], index: int) -> None:
    assert run[0][index] == RECORD['cases'][index][1]


@pytest.mark.parametrize('index', range(len(RECORD['cases'])))
def test_each_package_as_excel_saved_it(run: tuple[list[str], Path], index: int) -> None:
    assert saved(run[1] / f'c{index}.xlsx', red_png()) == RECORD['cases'][index][2]
