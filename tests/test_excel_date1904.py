"""Workbook date systems against independent Excel runtime and persisted controls."""
from __future__ import annotations

import hashlib
import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

ROOT = Path(__file__).parent / 'fixtures/date1904'
RAW: dict[str, Any] = json.loads((ROOT / 'native.json').read_bytes())
MANIFEST: dict[str, str] = json.loads((ROOT / 'manifest.json').read_bytes())


def test_native_evidence_exact() -> None:
    assert RAW['completed'] and RAW['closed_utc']
    for name, expected in MANIFEST.items():
        assert hashlib.sha256((ROOT / name).read_bytes()).hexdigest() == expected


@pytest.mark.parametrize('control', RAW['probes'], ids=lambda control: control['name'])
def test_native_runtime(control: dict[str, Any]) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(control['source'], name='DateControl')
    assert app.run('Probe') == control['result']


@pytest.mark.parametrize('control', RAW['probes'], ids=lambda control: control['name'])
def test_native_import_and_generated_roundtrip(control: dict[str, Any], tmp_path: Path) -> None:
    path = tmp_path / (control['name'] + '.xlsx')
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(control['source'], name='DateControl')
    app.run('Probe')
    app.save(path)
    observer = control['source'].split('Function Probe()')[0]
    for selected in (ROOT / path.name, path):
        loaded = ExcelApplication.open(selected)
        loaded.add_module(observer, name='InspectControl')
        assert loaded.run('Snapshot') == control['reopened_result']


@pytest.mark.parametrize('prefix', ['', 's:'])
def test_property_patch_preserves_opaque_metadata(prefix: str) -> None:
    from pyopenvba.apps.excel._date_system import NS, read, write

    app = ExcelApplication()
    app.add_workbook()
    declaration = 'xmlns:s' if prefix else 'xmlns'
    raw = (f'<{prefix}workbook {declaration}="{NS}"><{prefix}fileVersion/>'
           f'<{prefix}workbookPr other="contains date1904=0 and &gt;" date1904="false"/>'
           f'<{prefix}bookViews/><{prefix}sheets/></{prefix}workbook>').encode()
    book = app.workbook
    read(book, raw)
    assert not book.epoch_1904
    book.set_Date1904(True)
    result = write(book, raw)
    assert result == raw.replace(b'date1904="false"', b'date1904="1"')
    read(book, result)
    assert book.epoch_1904


@pytest.mark.parametrize('body', ['<workbookPr date1904="bad"/>', '<workbookPr/><workbookPr/>'])
def test_invalid_properties_are_refused(body: str) -> None:
    from pyopenvba.apps.excel._date_system import NS, read
    from pyopenvba.exceptions import VBAUnsupportedError

    app = ExcelApplication()
    app.add_workbook()
    raw = f'<workbook xmlns="{NS}">{body}</workbook>'.encode()
    with pytest.raises(VBAUnsupportedError):
        read(app.workbook, raw)


def test_empty_property_insertion_and_unmodified_bytes() -> None:
    from pyopenvba.apps.excel._date_system import NS, read, write

    app = ExcelApplication()
    app.add_workbook()
    raw = f'<workbook xmlns="{NS}"><fileVersion/><sheets/></workbook>'.encode()
    read(app.workbook, raw)
    assert write(app.workbook, raw) is raw
    app.workbook.set_Date1904(True)
    changed = write(app.workbook, raw)
    assert b'<fileVersion/><workbookPr date1904="1"/><sheets/>' in changed
