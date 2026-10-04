"""Runtime sources survive package saves and execute after reopening."""
from __future__ import annotations

import importlib
import os
from pathlib import Path
import zipfile

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.excel import ExcelFile
from pyopenvba.exceptions import VBAProjectError


def project_bytes(path: Path) -> bytes:
    with zipfile.ZipFile(path) as package:
        return package.read("xl/vbaProject.bin")


def make_runtime(app: ExcelApplication | None = None) -> ExcelApplication:
    if app is None:
        app = ExcelApplication()
        app.add_workbook()
    app.add_module('Public Function Twice(ByVal n As Long) As Long\nTwice = n * 2\nEnd Function', name="Multiplier", kind="class")
    app.add_module('Public Function SheetAnswer() As Long\nSheetAnswer = 3\nEnd Function', name="Sheet1", kind="document")
    app.add_module('Public Function BookAnswer() As Long\nBookAnswer = 4\nEnd Function', name="ThisWorkbook", kind="document")
    app.add_module('Public Function Answer() As Long\nDim item As New Multiplier\nAnswer = item.Twice(21) + Sheet1.SheetAnswer + ThisWorkbook.BookAnswer\nEnd Function', name="RuntimeCode")
    return app


def test_runtime_project_roundtrip_and_unchanged_save(tmp_path: Path) -> None:
    app = make_runtime()
    path = app.save(tmp_path / "runtime.xlsm")
    reopened = ExcelApplication.open(path)
    assert reopened.run("Answer") == 49
    second = reopened.save(tmp_path / "unchanged.xlsm")
    assert project_bytes(second) == project_bytes(path)
    reopened.add_module('Public Function Answer() As Long\nAnswer = 123\nEnd Function', name="RuntimeCode")
    edited = reopened.save(tmp_path / "edited.xlsm")
    assert ExcelApplication.open(edited).run("Answer") == 123
    with ExcelFile(edited) as host:
        assert 'Attribute VB_Name = "RuntimeCode"' in host.vba_modules()["RuntimeCode"]


def test_runtime_project_is_not_copied_into_another_workbook(tmp_path: Path) -> None:
    app = make_runtime()
    other = app.add_workbook()
    app.activate_workbook(other)
    path = app.save(tmp_path / "other.xlsm")
    with ExcelFile(path) as host:
        assert "RuntimeCode" not in host.vba_modules()


def test_macro_free_save_does_not_persist_runtime_source(tmp_path: Path) -> None:
    path = make_runtime().save(tmp_path / "macro-free.xlsx")
    with zipfile.ZipFile(path) as package:
        assert "xl/vbaProject.bin" not in package.namelist()


def test_runtime_creates_project_in_macro_free_workbook(tmp_path: Path) -> None:
    original = make_runtime().save(tmp_path / "macro-free.xlsx")
    app = ExcelApplication.open(original)
    app.add_module('Public Function RuntimeAnswer() As Long\nRuntimeAnswer = 42\nEnd Function', name="RuntimeCode")
    path = app.save(tmp_path / "new-project.xlsm")
    assert ExcelApplication.open(path).run("RuntimeAnswer") == 42
    with ExcelFile(path) as host:
        assert {"ThisWorkbook", "Sheet1", "RuntimeCode"} <= set(host.vba_modules())


def test_signed_project_is_preserved_until_source_changes(tmp_path: Path) -> None:
    source = Path(__file__).parent / "fixtures/signature_parts/excel.xlsm"
    app = ExcelApplication.open(source)
    unchanged = app.save(tmp_path / "unchanged.xlsm")
    assert project_bytes(unchanged) == project_bytes(source)
    with ExcelFile(unchanged) as host:
        assert host.vba_signature().present
    app.add_module('Public Function RuntimeAnswer() As Long\nRuntimeAnswer = 42\nEnd Function', name="RuntimeCode")
    with pytest.warns(UserWarning, match="signature"):
        edited = app.save(tmp_path / "edited.xlsm")
    with ExcelFile(edited) as host:
        assert not host.vba_signature().present
    assert ExcelApplication.open(edited).run("RuntimeAnswer") == 42


def test_protected_project_refuses_runtime_source_changes(tmp_path: Path) -> None:
    source = Path(__file__).parent / "live_excel_testing/workbook_with_password_protected_vba_modules.xlsm"
    app = ExcelApplication.open(source)
    app.add_module('Public Function RuntimeAnswer() As Long\nRuntimeAnswer = 42\nEnd Function', name="RuntimeCode")
    target = tmp_path / "protected.xlsm"
    with pytest.raises(VBAProjectError, match="password-protected"):
        app.save(target)
    assert not target.exists()


def test_unchanged_userform_designer_survives_source_edit(tmp_path: Path) -> None:
    source = Path(__file__).parent / "live_excel_testing/nested_form.xlsm"
    app = ExcelApplication.open(source)
    app.add_module('Public Function RuntimeAnswer() As Long\nRuntimeAnswer = 42\nEnd Function', name="RuntimeCode")
    path = app.save(tmp_path / "forms.xlsm")
    with ExcelFile(source) as original, ExcelFile(path) as edited:
        assert [form.name for form in edited.forms()] == [form.name for form in original.forms()]
    from pyopenvba.cfb import CFB

    before, after = CFB.from_bytes(project_bytes(source)), CFB.from_bytes(project_bytes(path))
    for name in before.list_streams():
        if name.startswith("FrmNested/"):
            assert after.get_stream(name) == before.get_stream(name)


@pytest.mark.skipif(os.environ.get("RUN_LIVE_EXCEL") != "1", reason="requires isolated real Excel")
@pytest.mark.parametrize("new_project", [False, True])
def test_native_excel_executes_runtime_project(tmp_path: Path, new_project: bool) -> None:
    harness = importlib.import_module("pyvbaharness")
    app = make_runtime()
    if new_project:
        app = make_runtime(ExcelApplication.open(app.save(tmp_path / "macro-free.xlsx")))
    path = app.save(tmp_path / "runtime.xlsm")
    with harness.ExcelSession(harness.HarnessConfig(lock_wait_s=30.0)) as excel:
        excel.new_document()
        source = ('Function Probe() As Long\nDim book As Workbook\nSet book = Workbooks.Open("'
                  + str(path).replace('"', '""') + '")\nProbe = Application.Run("\'" & book.Name & "\'!Answer")\nbook.Close False\nEnd Function')
        result = excel.run_vba(source, "Probe", timeout=30.0)
        assert result.ok, result
        assert not result.dialogs
        assert result.value == 49
