"""Independent Excel VBA projects, measured lookup and saved state."""
from __future__ import annotations

import json
import importlib
import os
from pathlib import Path

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.excel import ExcelFile
from pyopenvba.exceptions import VBACompileError, VBARuntimeError

MEASURED = json.loads((Path(__file__).parent / "fixtures/excel_projects.json").read_text(encoding="utf-8"))


@pytest.mark.parametrize("probe", MEASURED["probes"], ids=lambda probe: probe["name"])
def test_excel_project_lookup(probe: dict[str, str], tmp_path: Path) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.save(tmp_path / "caller.xlsm")
    for name in ("alpha", "beta"):
        book = app.add_workbook()
        app.add_module(MEASURED["module"], name="Library", workbook=book)
        app.add_module(MEASURED["class"], name="Box", kind="class", workbook=book)
        app.add_module(MEASURED["document"], name="ThisWorkbook", kind="document", workbook=book)
        book.SaveAs(str(tmp_path / f"{name}.xlsm"), 52)
    caller_source = probe["caller_source"]
    source = ('Function Probe() As String\nDim alpha As Workbook, beta As Workbook\n'
              'Dim answer As Variant, number As Long\n'
              'Set alpha = Workbooks("alpha.xlsm")\nSet beta = Workbooks("beta.xlsm")\n'
              'On Error Resume Next\n' + probe["action"]
              + '\nnumber = Err.Number\nOn Error GoTo 0\nProbe = CStr(answer) & "#error:" & CStr(number)\nEnd Function\n'
              + caller_source)
    app.add_module(source, name="ProbeModule")
    assert app.run("Probe") == probe["result"]
    assert not app.application.executing_projects


def test_opened_projects_keep_globals_documents_and_saved_source_separate(tmp_path: Path) -> None:
    paths: list[Path] = []
    for name, increment in (("alpha", 1), ("beta", 10)):
        app = ExcelApplication()
        app.add_workbook()
        app.add_module(f'Public Count As Long\nPublic Function NextValue() As Long\nCount = Count + {increment}\nNextValue = Count\nEnd Function\nPublic Function DocumentIdentity() As String\nDocumentIdentity = Sheet1.Identity\nEnd Function', name="Counter")
        app.add_module('Public Function Identity() As String\nIdentity = Me.Name & ":" & ThisWorkbook.Name\nEnd Function', name="Sheet1", kind="document")
        paths.append(app.save(tmp_path / f"{name}.xlsm"))
    app = ExcelApplication.open(paths[0])
    other = app.open_workbook(paths[1])
    assert app.run("'alpha.xlsm'!NextValue") == 1
    assert app.run("'beta.xlsm'!NextValue") == 10
    assert app.run("'alpha.xlsm'!NextValue") == 2
    assert app.run("'beta.xlsm'!NextValue") == 20
    assert app.run("'alpha.xlsm'!DocumentIdentity") == "Sheet1:alpha.xlsm"
    assert app.run("'beta.xlsm'!DocumentIdentity") == "Sheet1:beta.xlsm"
    app.add_module('Public Function NextValue() As Long\nNextValue = 99\nEnd Function', name="Counter", workbook=other)
    app.save(paths[1])
    assert ExcelApplication.open(paths[1]).run("NextValue") == 99
    assert ExcelApplication.open(paths[0]).run("NextValue") == 1


def test_nested_project_error_restores_the_caller(tmp_path: Path) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.save(tmp_path / "caller.xlsm")
    other = app.add_workbook()
    app.add_module('Public Function Fail() As Long\nErr.Raise 5\nEnd Function', name="Failing", workbook=other)
    other.SaveAs(str(tmp_path / "other.xlsm"), 52)
    app.add_module('Public Function Probe() As String\nOn Error Resume Next\nApplication.Run "\'other.xlsm\'!Fail"\nProbe = CStr(Err.Number) & ":" & ThisWorkbook.Name & ":" & Application.ThisWorkbook.Name\nEnd Function', name="ProbeModule")
    assert app.run("Probe") == "5:caller.xlsm:caller.xlsm"
    assert not app.application.executing_projects


def test_invalid_project_opens_as_a_complete_editable_project(tmp_path: Path) -> None:
    path = tmp_path / "broken.xlsm"
    with ExcelFile.create_new(path) as host:
        host.set_module("Module1", 'Public Function ValidAnswer() As Long\nValidAnswer = 42\nEnd Function')
        host.vba_project().add_module("BrokenModule", 'Public Function Broken(\nEnd Function')
        host.save()
    app = ExcelApplication()
    original = app.add_workbook()
    opened = app.open_workbook(path)
    assert app.workbooks() == [original, opened]
    assert app.workbook is opened
    from pyopenvba.apps.excel._vbide import VBProject

    project = opened.VBProject()
    assert isinstance(project, VBProject)
    assert {'module1', 'brokenmodule'} <= {entry.name.casefold() for entry in project.components.entries}
    with pytest.raises(VBACompileError):
        app.run("'broken.xlsm'!ValidAnswer")
    blank = ExcelApplication()
    other = blank.open_workbook(path)
    assert blank.workbooks() == [other]
    assert blank.interpreter.bridge.workbook is other


def test_closed_workbook_cannot_be_called_by_name(tmp_path: Path) -> None:
    app = ExcelApplication()
    book = app.add_workbook()
    app.add_module('Public Function Answer() As Long\nAnswer = 42\nEnd Function', name="Code")
    path = app.save(tmp_path / "closed.xlsm")
    book.Close(SaveChanges=False)
    with pytest.raises(VBARuntimeError) as failure:
        app.run(f"'{path.name}'!Answer")
    assert failure.value.number == 1004


def test_invalid_document_name_does_not_modify_project(tmp_path: Path) -> None:
    app = ExcelApplication()
    book = app.add_workbook()
    app.save(tmp_path / "original.xlsm")
    with pytest.raises(ValueError, match="code name"):
        app.add_module('Public Function Answer() As Long\nAnswer = 42\nEnd Function', name="AbsentSheet", kind="document")
    assert book.saved
    with pytest.raises(VBACompileError):
        app.interpreter.module("AbsentSheet")


OPEN = json.loads((Path(__file__).parent / "fixtures/workbook_open.json").read_text(encoding="utf-8"))


@pytest.mark.parametrize("probe", OPEN["probes"], ids=lambda probe: probe["name"])
def test_workbook_open_events(probe: dict[str, str], tmp_path: Path) -> None:
    target = ExcelApplication()
    book = target.add_workbook()
    target.add_module(OPEN["book"], name=book.code_name, kind="document")
    target.add_module(OPEN["sheet"], name="Sheet1", kind="document")
    target_path = target.save(tmp_path / "target.xlsm")
    app = ExcelApplication()
    app.add_workbook()
    app.save(tmp_path / "caller.xlsm")
    app.add_module(OPEN["log"], name="EventLogModule")
    app.add_module(OPEN["sink"], name="OpenSink", kind="class")
    app.add_module(probe["source"].replace("{TARGET}", str(target_path)), name="ProbeModule")
    assert app.run("Probe") == probe["result"]
    assert not app.application.executing_projects


WATCHER = '''Public WithEvents Source As Worksheet
Private Sub Source_Change(ByVal Target As Range)
Seen = Seen & ThisWorkbook.Name & "|"
End Sub
'''
WATCH_STATE = '''Public Watcher As Listener
Public Seen As String
Public Sub Bind(ByVal source As Worksheet)
Set Watcher = New Listener
Set Watcher.Source = source
End Sub
Public Sub Disconnect()
Set Watcher.Source = Nothing
End Sub
Public Function ReadSeen() As String
ReadSeen = Seen
End Function
'''


def watcher_files(tmp_path: Path) -> tuple[Path, Path]:
    paths: list[Path] = []
    for name in ("caller", "alpha"):
        app = ExcelApplication()
        app.add_workbook()
        app.add_module(WATCHER, name="Listener", kind="class")
        app.add_module(WATCH_STATE, name="State")
        paths.append(app.save(tmp_path / f"{name}.xlsm"))
    return paths[0], paths[1]


def test_shared_events_execute_in_each_subscribers_project(tmp_path: Path) -> None:
    caller, alpha = watcher_files(tmp_path)
    app = ExcelApplication.open(caller)
    app.open_workbook(alpha)
    target = app.add_workbook()
    sheet = target.sheets_[0]
    app.run("Bind", sheet)
    app.run("'alpha.xlsm'!Bind", sheet)
    app.sheet(1).set_value("A1", 1)
    app.run("Disconnect")
    app.sheet(1).set_value("A1", 2)
    assert app.run("ReadSeen") == "caller.xlsm|"
    assert app.run("'alpha.xlsm'!ReadSeen") == "alpha.xlsm|alpha.xlsm|"


@pytest.mark.skipif(os.environ.get("RUN_LIVE_EXCEL") != "1", reason="requires isolated real Excel")
def test_native_excel_executes_persisted_project_subscriptions(tmp_path: Path) -> None:
    harness = importlib.import_module("pyvbaharness")
    caller, alpha = watcher_files(tmp_path)
    source = ('Function Probe() As String\nDim caller As Workbook, alpha As Workbook, target As Workbook\n'
              'Application.EnableEvents = False\nSet caller = Workbooks.Open("' + str(caller).replace('"', '""') + '")\n'
              'Set alpha = Workbooks.Open("' + str(alpha).replace('"', '""') + '")\n'
              'Set target = Workbooks.Add\nApplication.EnableEvents = True\n'
              'Application.Run "\'caller.xlsm\'!Bind", target.Worksheets(1)\n'
              'Application.Run "\'alpha.xlsm\'!Bind", target.Worksheets(1)\n'
              'target.Worksheets(1).Range("A1").Value = 1\n'
              'Application.Run "\'caller.xlsm\'!Disconnect"\n'
              'target.Worksheets(1).Range("A1").Value = 2\n'
              'Probe = Application.Run("\'caller.xlsm\'!ReadSeen") & ";" & Application.Run("\'alpha.xlsm\'!ReadSeen")\n'
              'Application.EnableEvents = False\ntarget.Close False\nalpha.Close False\ncaller.Close False\nEnd Function')
    with harness.ExcelSession(harness.HarnessConfig(lock_wait_s=30.0)) as excel:
        excel.new_document()
        result = excel.run_vba(source, "Probe", timeout=30.0)
        assert result.ok, result
        assert not result.dialogs
        assert result.value == "caller.xlsm|;alpha.xlsm|alpha.xlsm|"


@pytest.mark.skipif(os.environ.get("RUN_LIVE_EXCEL") != "1", reason="requires isolated real Excel")
def test_native_excel_executes_persisted_open_handlers(tmp_path: Path) -> None:
    harness = importlib.import_module("pyvbaharness")
    target = ExcelApplication()
    book = target.add_workbook()
    target.add_module(OPEN["book"], name=book.code_name, kind="document")
    target.add_module(OPEN["sheet"], name="Sheet1", kind="document")
    path = target.save(tmp_path / "target.xlsm")
    source = ('Function Probe() As String\nDim target As Workbook\nApplication.EnableEvents = False\n'
              'Application.DisplayAlerts = False\nThisWorkbook.SaveAs "' + str(tmp_path / "caller.xlsm").replace('"', '""') + '", 52\n'
              'EventLog = ""\nApplication.EnableEvents = True\nSet target = Workbooks.Open("' + str(path).replace('"', '""') + '")\n'
              'Probe = EventLog & "#value:" & CStr(target.Worksheets(1).Range("A1").Value)\n'
              'Application.EnableEvents = False\ntarget.Close False\nEnd Function')
    with harness.ExcelSession(harness.HarnessConfig(lock_wait_s=30.0)) as excel:
        excel.new_document()
        excel.add_module("EventLogModule", OPEN["log"])
        result = excel.run_vba(source, "Probe", timeout=30.0)
        assert result.ok, result
        assert not result.dialogs
        assert result.value == ('book-open:target.xlsm:target.xlsm|sheet-change:Sheet1|book-change:Sheet1|'
                                'book-activate:target.xlsm|#value:41')
