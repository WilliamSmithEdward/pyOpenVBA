"""Native post-Apply/imported table-sort configuration persistence."""
from __future__ import annotations

import json
from hashlib import sha256
from pathlib import Path, PureWindowsPath
from typing import Any
from zipfile import ZipFile

import pytest

from pyopenvba.apps.excel import ExcelApplication

ROOT = Path(__file__).parent / 'fixtures/table_sort/post_apply'
RECORD: dict[str, Any] = json.loads((ROOT / 'native.json').read_text(encoding='utf8'))
MANIFEST: dict[str, Any] = json.loads((ROOT / 'manifest.json').read_text(encoding='utf8'))


def configured(case: dict[str, Any]) -> ExcelApplication:
    if case['loaded']:
        sample = ROOT.parent / PureWindowsPath(case['input_file']).name
        app = ExcelApplication.open(sample, with_vba=False)
    else:
        app = ExcelApplication()
        app.add_workbook()
        app.add_module(case['setup_source'], name='Setup')
        assert app.run('Setup.Probe') == case['setup']['value']
    app.add_module(case['inspection_source'], name='Before')
    assert app.run('Before.Snapshot') == case['before']['value']
    app.add_module(case['source'], name='Mutation')
    return app


@pytest.mark.parametrize('case', RECORD['probes'], ids=lambda case: case['name'])
def test_native_table_sort_live_configuration_save(case: dict[str, Any], tmp_path: Path) -> None:
    app = configured(case)
    assert app.run('Mutation.Probe') == case['result']['value']
    saved = app.save(tmp_path / 'table-sort-options.xlsm')
    reopened = ExcelApplication.open(saved, with_vba=False)
    reopened.add_module(case['inspection_source'], name='Inspect')
    assert reopened.run('Snapshot') == case['reopened']['value']


@pytest.mark.parametrize('case', RECORD['probes'], ids=lambda case: case['name'])
def test_original_post_apply_native_input(case: dict[str, Any]) -> None:
    app = ExcelApplication.open(ROOT / PureWindowsPath(case['file']).name, with_vba=False)
    app.add_module(case['inspection_source'], name='Inspect')
    assert app.run('Snapshot') == case['reopened']['value']


def test_post_apply_raw_package_provenance() -> None:
    assert RECORD['completed'] and RECORD['closed_utc']
    assert sha256((ROOT / 'native.json').read_bytes()).hexdigest() == MANIFEST['raw_sha256']
    for case in RECORD['probes']:
        for phase in ('before', 'result', 'save_status', 'reopened'):
            assert case[phase]['ok'] and not case[phase]['dialogs']
        entry = MANIFEST['cases'][case['name']]
        assert sha256(case['source'].encode()).hexdigest() == entry['source_sha256']
        assert sha256(case['result']['value'].encode()).hexdigest() == entry['result_sha256']
        assert sha256(case['reopened']['value'].encode()).hexdigest() == entry['reopened_sha256']
        sample = ROOT / PureWindowsPath(case['file']).name
        assert sha256(sample.read_bytes()).hexdigest() == MANIFEST['samples'][sample.name]
        with ZipFile(sample) as package:
            assert not any('vbaproject' in name.casefold() for name in package.namelist())
            for name, expected in case['xml'].items():
                assert package.read(name).decode('utf8') == expected
