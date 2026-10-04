"""Native-recorded VBIDE file operations and physical source conversions."""
from __future__ import annotations

import importlib
import json
import os
from pathlib import Path

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.apps.excel._vbide import VBComponent, VBProject
from pyopenvba.excel import ExcelFile

MEASURED = json.loads((Path(__file__).parent / 'fixtures/vbide_source_files.json').read_text(encoding='utf-8'))


@pytest.mark.parametrize('probe', MEASURED['probes'], ids=[probe['name'] for probe in MEASURED['probes']])
def test_native_source_file_operations(probe: dict[str, str | None], tmp_path: Path) -> None:
    name = str(probe['name'])
    input_path = tmp_path / ('input.cls' if name.endswith('class') else 'input.bas')
    input_path.write_bytes(MEASURED['cls' if name.endswith('class') else 'bas'].encode('cp1252'))
    output_path = tmp_path / 'output.txt'
    source = str(probe['macro']).replace('"INPUT"', '"' + str(input_path).replace('"', '""') + '"').replace('"OUTPUT"', '"' + str(output_path).replace('"', '""') + '"')
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(source, name='Harness')
    result = app.run('Harness.Probe')
    assert isinstance(result, str)
    assert result.replace('\r\n', '\n') == probe['result']
    if probe['exported'] is not None:
        assert output_path.read_bytes().decode('cp1252') == probe['exported']
    else:
        assert not output_path.exists()


def imported_workbook(tmp_path: Path) -> Path:
    app = ExcelApplication()
    app.add_workbook()
    project = app.workbook.VBProject()
    assert isinstance(project, VBProject)
    bas = tmp_path / 'input.bas'
    cls = tmp_path / 'input.cls'
    bas.write_bytes(MEASURED['bas'].encode('cp1252'))
    cls.write_bytes(MEASURED['cls'].encode('cp1252'))
    project.components.Import(str(bas))
    project.components.Import(str(cls))
    app.add_module('Function Combined() As Long\nDim value As New ImportedClass\nCombined = Imported.Answer() + value.Answer()\nEnd Function', name='RunImports')
    assert app.run('RunImports.Combined') == 84
    return app.save(tmp_path / 'imports.xlsm')


def test_imported_standard_and_class_sources_persist_and_execute(tmp_path: Path) -> None:
    path = imported_workbook(tmp_path)
    reopened = ExcelApplication.open(path)
    assert reopened.run('RunImports.Combined') == 84
    with ExcelFile(path) as host:
        assert 'VERSION 1.0 CLASS' not in host.get_module('ImportedClass')
        assert 'Attribute VB_Base' in host.get_module('ImportedClass')


def test_python_source_edit_refreshes_existing_editor_buffer() -> None:
    app = ExcelApplication()
    app.add_workbook()
    project = app.workbook.VBProject()
    assert isinstance(project, VBProject)
    component = project.components.Add(1)
    assert isinstance(component, VBComponent)
    component.code_module.AddFromString('Function Answer() As Long\nAnswer = 11\nEnd Function')
    app.add_module('Function Answer() As Long\nAnswer = 29\nEnd Function', name='Module1')
    assert 'Answer = 29' in str(component.code_module.Lines(1, 3))
    assert app.run('Module1.Answer') == 29


@pytest.mark.skipif(os.environ.get('RUN_LIVE_EXCEL') != '1', reason='requires isolated real Excel')
def test_native_excel_executes_imported_standard_and_class_code(tmp_path: Path) -> None:
    harness = importlib.import_module('pyvbaharness')
    path = imported_workbook(tmp_path)
    source = ('Function Probe() As Long\nDim target As Workbook\nApplication.EnableEvents = False\n'
              'Set target = Workbooks.Open("' + str(path).replace('"', '""') + '")\n'
              'Probe = Application.Run("\'imports.xlsm\'!RunImports.Combined")\n'
              'target.Close False\nEnd Function')
    with harness.ExcelSession(harness.HarnessConfig(lock_wait_s=30.0)) as excel:
        excel.new_document()
        result = excel.run_vba(source, 'Probe', timeout=30.0)
        assert result.ok and not result.dialogs, result
        assert result.value == 84
