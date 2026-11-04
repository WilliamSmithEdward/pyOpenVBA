"""Print settings read from a file Excel saved, and from the model's own save of that file (scripts/measure_page_files.py,
tests/fixtures/page_files.json): every property of every sheet, against what Excel read from the same file."""
from __future__ import annotations

import base64
import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/page_files.json').read_text(encoding='utf-8'))
READ_SOURCE = RECORD['reader']


def _read(path: Path) -> str:
    """Each sheet read in the model by the macro that read it in Excel."""
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(READ_SOURCE, name='Reader')
    return str(app.run('Reader.Opened', str(path), ','.join(RECORD['read'])))


def _differences(got: str) -> list[tuple[str, str, str, str]]:
    names = [*RECORD['read'], 'Even.LeftHeader', 'Even.CenterFooter', 'First.RightHeader', 'First.LeftFooter']
    found: list[tuple[str, str, str, str]] = []
    for ours, theirs in zip(got.split('@@@'), str(RECORD['answer']).split('@@@'), strict=True):
        sheet, *mine = ours.split('~')
        _, *excel = theirs.split('~')
        found += [(sheet, name, one, other) for name, one, other in zip(names, excel, mine, strict=True) if one != other]
    return found


@pytest.fixture
def excel_file(tmp_path: Path) -> Path:
    path = tmp_path / 'page_files.xlsx'
    path.write_bytes(base64.b64decode(RECORD['file']))
    return path


def test_what_excel_saved_reads_as_excel_read_it(excel_file: Path) -> None:
    assert not _differences(_read(excel_file))


def test_the_models_save_of_it_keeps_every_setting(excel_file: Path, tmp_path: Path) -> None:
    app = ExcelApplication.open(excel_file)
    saved = app.save(tmp_path / 'again.xlsx')
    assert not _differences(_read(Path(saved)))
