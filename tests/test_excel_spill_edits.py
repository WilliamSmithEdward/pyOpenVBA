"""Replay native structural edits and copies through dynamic-array spills."""
from __future__ import annotations

import json
import importlib.util
import os
from pathlib import Path

import pytest

from pyopenvba.apps.excel import ExcelApplication

PROBES: list[dict[str, str]] = json.loads((Path(__file__).parent / 'fixtures/spill_edits.json').read_text(encoding='utf-8'))['probes']
READ_SAVED = '''Function ReadSaved() As String
Dim ws As Worksheet, cell As Range
Set ws = ActiveSheet
ReadSaved = "0"
For Each cell In ws.Range("A1:F6")
ReadSaved = ReadSaved & "|" & cell.Address(False,False) & ":" & CStr(cell.Value2) & ":" & cell.Formula2 & ":" & cell.HasSpill
Next
End Function'''

@pytest.mark.parametrize('probe', PROBES, ids=[probe['name'] for probe in PROBES])
def test_native_spill_edit(probe: dict[str, str]) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(probe['macro'], name='Harness')
    assert app.run('Harness.Probe') == probe['result']


@pytest.mark.parametrize('probe', PROBES, ids=[probe['name'] for probe in PROBES])
def test_spill_edit_survives_save_reopen(probe: dict[str, str], tmp_path: Path) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(probe['macro'], name='Harness')
    app.run('Harness.Probe')
    reopened = ExcelApplication.open(app.save(tmp_path / 'edited_spill.xlsx'))
    reopened.add_module(READ_SAVED, name='Readback')
    assert reopened.run('ReadSaved') == '0|' + probe['result'].split('|', 1)[1]


@pytest.mark.skipif(os.environ.get('RUN_LIVE_EXCEL') != '1' or importlib.util.find_spec('pyvbaharness') is None,
                    reason='requires explicitly enabled native Excel')
def test_native_excel_opens_edited_spill(tmp_path: Path) -> None:
    harness = pytest.importorskip('pyvbaharness')
    saved: list[tuple[Path, str]] = []
    for name in ['insert_inside', 'insert_column_inside', 'copy_whole', 'copy_overlap']:
        probe = next(probe for probe in PROBES if probe['name'] == name)
        app = ExcelApplication()
        app.add_workbook()
        app.add_module(probe['macro'], name='Harness')
        app.run('Harness.Probe')
        saved.append((app.save(tmp_path / f'{name}.xlsx'), probe['result']))
    with harness.ExcelSession() as excel:
        for path, expected in saved:
            excel.open_document(path)
            result = excel.run_vba(READ_SAVED, 'ReadSaved', timeout=60.0)
            assert result.ok and not result.dialogs, result
            assert result.value == expected
