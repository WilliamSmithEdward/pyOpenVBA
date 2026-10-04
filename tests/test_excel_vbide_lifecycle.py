"""Native module resets and opening/editing incomplete VBA source."""
from __future__ import annotations

import json
from pathlib import Path

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.apps.excel._vbide import VBComponent, VBProject
from pyopenvba.exceptions import VBACompileError

MEASURED = json.loads((Path(__file__).parent / 'fixtures/vbide_compile_lifecycle.json').read_text(encoding='utf-8'))


@pytest.mark.parametrize('probe', MEASURED['probes'], ids=[probe['name'] for probe in MEASURED['probes']])
def test_native_vbide_lifecycle(probe: dict[str, str], tmp_path: Path) -> None:
    app = ExcelApplication()
    app.add_workbook()
    path = tmp_path / 'invalid.xlsm'
    source = probe['macro'].replace('"OUTPUT"', '"' + str(path).replace('"', '""') + '"')
    app.add_module(source, name='Harness')
    assert app.run('Harness.Probe') == probe['result']


def invalid_workbook(tmp_path: Path) -> Path:
    app = ExcelApplication()
    app.add_workbook()
    project = app.workbook.VBProject()
    assert isinstance(project, VBProject)
    component = project.components.Add(1)
    assert isinstance(component, VBComponent)
    component.code_module.AddFromString('this is invalid syntax')
    return app.save(tmp_path / 'invalid.xlsm')


def test_invalid_open_can_be_inspected_saved_repaired_and_executed(tmp_path: Path) -> None:
    path = invalid_workbook(tmp_path)
    app = ExcelApplication.open(path)
    assert app.workbook.saved
    project = app.workbook.VBProject()
    assert isinstance(project, VBProject)
    component = project.components.vba_get('Item', ['Module1'])
    assert isinstance(component, VBComponent)
    assert 'invalid syntax' in str(component.code_module.Lines(1, 1))
    with pytest.raises(VBACompileError):
        app.run('Module1.Answer')
    preserved = app.save(tmp_path / 'preserved.xlsm')
    assert preserved.exists()
    component.code_module.DeleteLines(1)
    component.code_module.AddFromString('Function Answer() As Long\nAnswer = 47\nEnd Function')
    assert app.run('Module1.Answer') == 47
    repaired = app.save(tmp_path / 'repaired.xlsm')
    assert ExcelApplication.open(repaired).run('Module1.Answer') == 47


def test_explicit_python_source_import_remains_atomic_and_strict(tmp_path: Path) -> None:
    path = invalid_workbook(tmp_path)
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('Function Original() As Long\nOriginal = 31\nEnd Function', name='Original')
    before = dict(app.interpreter.modules)
    with pytest.raises(VBACompileError):
        app.load_vba(path)
    assert app.interpreter.modules == before
    assert app.run('Original.Original') == 31
