"""Three old-collection gaps, replayed from live Excel (scripts/measure_old_collection_leftovers.py,
tests/fixtures/old_collection_leftovers.json): AutoScaleFont on a loaded button, Shadow on a new and a loaded picture, with what each save
changed tag by tag, and DrawingObjects.Visible, which reads the sheet's first form control's."""
from __future__ import annotations

import importlib.util
import json
import re
import struct
import zipfile
import zlib
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

FIXTURES = Path(__file__).parent / 'fixtures'
RECORD: dict[str, Any] = json.loads((FIXTURES / 'old_collection_leftovers.json').read_text(encoding='utf-8'))
_spec = importlib.util.spec_from_file_location('measure_loaded_drawing_edits', Path(__file__).parents[1] / 'scripts/measure_loaded_drawing_edits.py')
assert _spec is not None and _spec.loader is not None
_script = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_script)
changes: Any = getattr(_script, 'changes')
DRAWING = re.compile(r'xl/drawings/drawing\d+\.xml$')
BASES = {'button': 'button base', 'new picture': 'new picture base', 'loaded picture': 'loaded base'}


def _png() -> bytes:
    def chunk(kind: bytes, data: bytes) -> bytes:
        return struct.pack('>I', len(data)) + kind + data + struct.pack('>I', zlib.crc32(kind + data))

    return b'\x89PNG\r\n\x1a\n' + chunk(b'IHDR', struct.pack('>IIBBBBB', 1, 1, 8, 2, 0, 0, 0)) + chunk(b'IDAT', zlib.compress(b'\x00\xff\x00\x00')) + chunk(b'IEND', b'')


def _app(binding: str = 'Late') -> ExcelApplication:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(RECORD['source'].replace('{declared}', 'Excel.DrawingObjects' if binding == 'Typed' else 'Object'), name='Probe')
    return app


def _saved(which: str, folder: Path) -> tuple[str, dict[str, str]]:
    book, out, picture = folder / 'in.xlsx', folder / f'out_{which}.xlsx', folder / 'dot.png'
    book.write_bytes((FIXTURES / 'old_collections.xlsx').read_bytes())
    picture.write_bytes(_png())
    answer = str(_app().run('Probe.Row', str(book), str(out), str(picture), which))
    with zipfile.ZipFile(out) as package:
        return answer, {name: package.read(name).decode('utf-8') for name in package.namelist() if DRAWING.search(name)}


def _quiet(found: list[list[Any]]) -> list[list[Any]]:
    return [one for one in found if not any('a16:creationId' in piece for piece in one[1] + one[2])]


@pytest.mark.parametrize('case', RECORD['saves'], ids=[one['which'] for one in RECORD['saves']])
def test_each_save_as_excel_did(case: dict[str, Any], tmp_path: Path) -> None:
    answer, parts = _saved(case['which'], tmp_path)
    assert answer == case['excel']
    _, untouched = _saved(BASES[case['which']], tmp_path)
    model = {name: _quiet(changes(untouched[name], part)) for name, part in parts.items()}
    excel = {name: _quiet(case['parts'].get(name, [])) for name in parts}
    assert model == excel


@pytest.mark.parametrize('binding', ['Typed', 'Late'])
def test_drawing_objects_visible_reads_the_first_form_control(binding: str, tmp_path: Path) -> None:
    book = tmp_path / 'in.xlsx'
    book.write_bytes((FIXTURES / 'old_collections.xlsx').read_bytes())
    assert str(_app(binding).run('Probe.Visibles', str(book))) == RECORD['visibles'][binding]


@pytest.mark.parametrize('binding', ['Typed', 'Late'])
def test_drawing_objects_visible_on_a_made_book(binding: str) -> None:
    assert str(_app(binding).run('Probe.Made')) == RECORD['made'][binding]
