"""Native-recorded VBIDE operations run through the headless VBA engine."""
from __future__ import annotations

import json
import importlib
import os
import zipfile
from pathlib import Path

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBACompileError
from pyopenvba.exceptions import VBARuntimeError
from pyopenvba.excel import ExcelFile
from pyopenvba.apps.excel._vbide import VBProject, VBComponent, CodeModule

MEASURED = json.loads((Path(__file__).parent / 'fixtures/vbide_projects.json').read_text(encoding='utf-8'))


def literal(text: str) -> str:
    return ' & vbCrLf & '.join('"' + line.replace('"', '""') + '"' for line in text.split('\n'))


@pytest.mark.parametrize('probe', MEASURED['probes'], ids=[probe['name'] for probe in MEASURED['probes']])
def test_native_vbide_project_operations(probe: dict[str, str]) -> None:
    app = ExcelApplication()
    app.add_workbook()
    source = ('Function Probe() As String\nDim target As Workbook, project As Object, parts As Object\n'
              'Dim part As Object, code As Object, item As Object, number As Long, extra As String\n'
              'Application.EnableEvents = False\nApplication.DisplayAlerts = False\n'
              'Set target = Workbooks.Add\nSet project = target.VBProject\nSet parts = project.VBComponents\n')
    if probe['name'].startswith('source_'):
        source += 'Set part = parts.Add(1)\nSet code = part.CodeModule\ncode.AddFromString ' + literal(MEASURED['initial_source']) + '\n'
    source += ('On Error Resume Next\n' + probe['action'] + '\nnumber = Err.Number\nOn Error GoTo 0\n'
               'Probe = "#error:" & CStr(number) & "#extra:" & extra & "#project:" & project.Name\n'
               'For Each item In parts\nProbe = Probe & "|" & item.Name & ":" & CStr(item.Type) & ":" & CStr(item.CodeModule.CountOfLines)\nNext\n'
               'If Not code Is Nothing Then\nProbe = Probe & "#decl:" & CStr(code.CountOfDeclarationLines)\n'
               'If code.CountOfLines > 0 Then Probe = Probe & "#source:" & code.Lines(1, code.CountOfLines)\nEnd If\n'
               'Probe = Probe & "#sheet:" & target.Worksheets(1).CodeName & "#saved:" & CStr(target.Saved)\n'
               'target.Close False\nEnd Function')
    app.add_module(source, name='Harness')
    result = app.run('Harness.Probe')
    assert isinstance(result, str)
    assert result.replace('\r\n', '\n') == probe['result']


def test_edited_code_compiles_when_run_and_invalid_edits_remain_inspectable() -> None:
    app = ExcelApplication()
    book = app.add_workbook()
    project = book.VBProject()
    assert isinstance(project, VBProject)
    parts = project.components
    component = parts.vba_get('Add', [1])
    assert isinstance(component, VBComponent)
    code = component.vba_get('CodeModule')
    assert isinstance(code, CodeModule)
    code.vba_get('AddFromString', ['Function Answer() As Long\nAnswer = 42\nEnd Function'])
    assert app.run('Module1.Answer') == 42
    code.vba_get('ReplaceLine', [2, 'Answer = 71'])
    assert app.run('Module1.Answer') == 42  # Native Excel keeps compiled code until invalidated.
    code.vba_get('InsertLines', [1, 'this is invalid syntax'])
    assert 'this is invalid syntax' in str(code.vba_get('Lines', [1, 100]))
    with pytest.raises(VBACompileError):
        app.run('Module1.Answer')
    code.vba_get('DeleteLines', [1])
    assert app.run('Module1.Answer') == 71


def edited_workbook(tmp_path: Path) -> Path:
    app = ExcelApplication()
    book = app.add_workbook()
    project = book.VBProject()
    assert isinstance(project, VBProject)
    component = project.components.vba_get('Item', ['Sheet1'])
    assert isinstance(component, VBComponent)
    component.code_module.AddFromString('Public Sub WriteMarker()\nRange("B1").Value = 83\nEnd Sub')
    component.vba_set('Name', 'Info')
    project.vba_set('Name', 'EditedProject')
    standard = project.components.Add(1)
    assert isinstance(standard, VBComponent)
    standard.code_module.AddFromString('Function Answer() As Long\nAnswer = 42\nEnd Function')
    standard.vba_set('Name', 'Answers')
    return app.save(tmp_path / 'edited.xlsm')


def test_pending_source_and_project_renames_persist(tmp_path: Path) -> None:
    path = edited_workbook(tmp_path)
    with ExcelFile(path) as host:
        assert host.vba_project().name == 'EditedProject'
        assert 'Info' in host.module_names()
        assert 'Sheet1' not in host.module_names()
        assert 'Answers' in host.module_names()
        assert 'Answer = 42' in host.get_module('Answers')
    reopened = ExcelApplication.open(path)
    assert reopened.run('Answers.Answer') == 42
    reopened.run('Info.WriteMarker')
    assert reopened.workbook.sheets_[0].Range('B1').vba_get('Value') == 83


def test_document_can_reference_class_added_later_in_the_edit_set() -> None:
    app = ExcelApplication()
    app.add_workbook()
    project = app.workbook.VBProject()
    assert isinstance(project, VBProject)
    document = project.components.vba_get('Item', ['Sheet1'])
    assert isinstance(document, VBComponent)
    document.code_module.AddFromString('Private Value As New Class1\nFunction Read() As Long\nRead = Value.Answer()\nEnd Function')
    cls = project.components.Add(2)
    assert isinstance(cls, VBComponent)
    cls.code_module.AddFromString('Function Answer() As Long\nAnswer = 29\nEnd Function')
    assert app.run('Sheet1.Read') == 29


def test_invalid_pending_source_can_be_saved_without_compiling(tmp_path: Path) -> None:
    app = ExcelApplication()
    app.add_workbook()
    project = app.workbook.VBProject()
    assert isinstance(project, VBProject)
    component = project.components.Add(1)
    assert isinstance(component, VBComponent)
    component.code_module.AddFromString('this is invalid syntax')
    path = app.save(tmp_path / 'invalid.xlsm')
    with ExcelFile(path) as host:
        assert 'this is invalid syntax' in host.get_module('Module1')


def test_catalog_reads_preserve_signed_project_bytes(tmp_path: Path) -> None:
    source = Path(__file__).parent / 'fixtures/signature_parts/excel.xlsm'
    app = ExcelApplication.open(source)
    project = app.workbook.VBProject()
    assert isinstance(project, VBProject)
    assert project.components.vba_items()
    target = app.save(tmp_path / 'read-only.xlsm')
    with zipfile.ZipFile(source) as before, zipfile.ZipFile(target) as after:
        assert before.read('xl/vbaProject.bin') == after.read('xl/vbaProject.bin')
    with ExcelFile(target) as host:
        assert host.vba_signature().present


def test_protected_project_catalog_is_not_exposed() -> None:
    source = Path(__file__).parent / 'live_excel_testing/workbook_with_password_protected_vba_modules.xlsm'
    app = ExcelApplication.open(source)
    project = app.workbook.VBProject()
    assert isinstance(project, VBProject)
    assert project.Protection() == 1
    with pytest.raises(VBARuntimeError) as failure:
        project.VBComponents()
    assert failure.value.number == 50289


def test_existing_userform_component_retains_designer_type() -> None:
    app = ExcelApplication.open(Path(__file__).parent / 'live_excel_testing/nested_form.xlsm')
    project = app.workbook.VBProject()
    assert isinstance(project, VBProject)
    form = project.components.vba_get('Item', ['FrmNested'])
    assert isinstance(form, VBComponent)
    assert form.Type() == 3


def test_copy_pending_document_source_and_python_edits_override_staged_source(tmp_path: Path) -> None:
    app = ExcelApplication()
    book = app.add_workbook()
    project = book.VBProject()
    assert isinstance(project, VBProject)
    component = project.components.vba_get('Item', ['Sheet1'])
    assert isinstance(component, VBComponent)
    component.code_module.AddFromString('Sub Mark()\nRange("B1").Value = 11\nEnd Sub')
    original = book.sheets_[0]
    original.Copy(After=original)
    app.run('Sheet2.Mark')
    assert book.sheets_[1].Range('B1').vba_get('Value') == 11
    component.code_module.ReplaceLine(2, 'Range("B1").Value = 22')
    app.add_module('Sub Mark()\nRange("B1").Value = 33\nEnd Sub', name='Sheet1', kind='document')
    app.run('Sheet1.Mark')
    assert original.Range('B1').vba_get('Value') == 33
    path = app.save(tmp_path / 'pending-copy.xlsm')
    reopened = ExcelApplication.open(path)
    reopened.run('Sheet1.Mark')
    reopened.run('Sheet2.Mark')
    assert reopened.workbook.sheets_[0].Range('B1').vba_get('Value') == 33
    assert reopened.workbook.sheets_[1].Range('B1').vba_get('Value') == 11


@pytest.mark.skipif(os.environ.get('RUN_LIVE_EXCEL') != '1', reason='requires isolated real Excel')
def test_native_excel_executes_vbide_edits_and_reads_renamed_project(tmp_path: Path) -> None:
    harness = importlib.import_module('pyvbaharness')
    path = edited_workbook(tmp_path)
    source = ('Function Probe() As String\nDim target As Workbook\nApplication.EnableEvents = False\n'
              'Set target = Workbooks.Open("' + str(path).replace('"', '""') + '")\n'
              'Application.Run "\'edited.xlsm\'!Info.WriteMarker"\n'
              'Probe = target.VBProject.Name & ":" & target.Worksheets(1).CodeName & ":" & '
              'CStr(target.Worksheets(1).Range("B1").Value) & ":" & CStr(Application.Run("\'edited.xlsm\'!Answers.Answer"))\n'
              'target.Close False\nEnd Function')
    with harness.ExcelSession(harness.HarnessConfig(lock_wait_s=30.0)) as excel:
        excel.new_document()
        result = excel.run_vba(source, 'Probe', timeout=30.0)
        assert result.ok and not result.dialogs, result
        assert result.value == 'EditedProject:Info:83:42'
