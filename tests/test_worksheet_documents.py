"""Worksheet documents share the tab lifecycle and persist as native VBA."""
from __future__ import annotations

import importlib
import json
import os
from pathlib import Path

import pytest

from pyopenvba.apps.excel import ExcelApplication, Worksheet
from pyopenvba.apps.excel._documents import materialize
from pyopenvba.excel import ExcelFile
from pyopenvba.exceptions import VBARuntimeError

MEASURED = json.loads((Path(__file__).parent / "fixtures/worksheet_documents.json").read_text(encoding="utf-8"))


def test_document_name_defaults_respect_source_attributes() -> None:
    app = ExcelApplication()
    book = app.add_workbook()
    runtime = app.add_module('Attribute VB_Name = "Sheet1"\n' + MEASURED["document"], kind="document")
    assert runtime.name == "Sheet1"
    assert book.sheets_[0].vba_document is not None
    workbook_module = app.add_module("Public Value As Long", kind="document")
    assert workbook_module.name == "ThisWorkbook"


@pytest.mark.parametrize("count,names,returned", [(1, ["Sheet2", "Sheet1"], "Sheet2"),
                                                (2, ["Sheet3", "Sheet2", "Sheet1"], "Sheet3"),
                                                (1.5, ["Sheet2", "Sheet1"], "Sheet2"),
                                                (2.9, ["Sheet3", "Sheet2", "Sheet1"], "Sheet3")])
def test_multiple_sheet_creation_and_lazy_names(count: float, names: list[str], returned: str) -> None:
    app = ExcelApplication()
    book = app.add_workbook()
    assert book.saved
    assert book.code_name == ""
    assert book.sheets_[0].code_name == ""
    app.add_module(MEASURED["document"], name="Sheet1", kind="document")
    added = book.Worksheets().vba_get("Add", [], {"Count": count})
    assert isinstance(added, Worksheet)
    assert [sheet.name for sheet in book.sheets_] == names
    assert added.name == returned
    assert added.code_name == ""
    materialize(book)
    assert added.code_name == "Sheet2"


@pytest.mark.parametrize("count", [0, -1])
def test_invalid_sheet_count_does_not_change_workbook(count: int) -> None:
    app = ExcelApplication()
    book = app.add_workbook()
    with pytest.raises(VBARuntimeError) as failure:
        book.Worksheets().vba_get("Add", [], {"Count": count})
    assert failure.value.number == 1004
    assert len(book.sheets_) == 1


def test_deleted_unmaterialized_sheet_does_not_consume_a_code_name() -> None:
    app = ExcelApplication()
    book = app.add_workbook()
    app.add_module(MEASURED["document"], name="Sheet1", kind="document")
    added = book.Worksheets().vba_get("Add")
    assert isinstance(added, Worksheet)
    added.Delete()
    replacement = book.Worksheets().vba_get("Add")
    assert isinstance(replacement, Worksheet)
    assert replacement.name == "Sheet3"
    materialize(book)
    assert replacement.code_name == "Sheet2"


def test_copy_resets_document_fields_and_persists_modules(tmp_path: Path) -> None:
    app = ExcelApplication()
    book = app.add_workbook()
    app.add_module(MEASURED["document"], name="Sheet1", kind="document")
    original = book.sheets_[0]
    app.run("Sheet1.Mark", 42)
    app.run("Sheet1.ExportMarker")
    original.Copy(After=original)
    copied = book.sheets_[1]
    assert copied.code_name == "Sheet2"
    app.run("Sheet2.ExportMarker")
    assert original.Range("B1").vba_get("Value") == 42
    assert copied.Range("B1").vba_get("Value") == 0
    path = app.save(tmp_path / "copied.xlsm")
    reopened = ExcelApplication.open(path)
    reopened.run("Sheet2.Mark", 7)
    reopened.run("Sheet2.ExportMarker")
    assert reopened.workbook.sheets_[1].Range("B1").vba_get("Value") == 7
    copied.Delete()
    app.save(path)
    with ExcelFile(path) as host:
        assert "Sheet2" not in host.vba_modules()
        assert "Sheet1" in host.vba_modules()


def test_copy_to_another_project_keeps_source_code_name_until_materialization(tmp_path: Path) -> None:
    app = ExcelApplication()
    origin = app.add_workbook()
    app.add_module(MEASURED["document"], name="Sheet1", kind="document")
    target = app.add_workbook()
    origin.sheets_[0].Copy(After=target.sheets_[0])
    copied = target.sheets_[1]
    assert copied.code_name == "Sheet1"
    assert target.code_name == ""
    assert target.sheets_[0].code_name == ""
    materialize(target)
    assert target.sheets_[0].code_name == "Sheet2"
    target.SaveAs(str(tmp_path / "target.xlsm"), 52)
    reopened = ExcelApplication.open(tmp_path / "target.xlsm")
    reopened.run("Sheet1.Mark", 17)
    reopened.run("Sheet1.ExportMarker")
    assert reopened.workbook.sheets_[1].Range("B1").vba_get("Value") == 17


def test_move_to_another_project_resets_fields_and_removes_source_module(tmp_path: Path) -> None:
    app = ExcelApplication()
    origin = app.add_workbook()
    app.add_module(MEASURED["document"], name="Sheet1", kind="document")
    origin.add_sheet("Keep")
    app.run("Sheet1.Mark", 42)
    target = app.add_workbook()
    origin.sheets_[0].Move(After=target.sheets_[0])
    assert [sheet.name for sheet in origin.sheets_] == ["Keep"]
    assert origin.project_runtime is not None
    assert "sheet1" not in origin.project_runtime.modules
    app.run(f"'{target.name}'!Sheet1.ExportMarker")
    assert target.sheets_[1].Range("B1").vba_get("Value") == 0
    target.SaveAs(str(tmp_path / "moved.xlsm"), 52)
    reopened = ExcelApplication.open(tmp_path / "moved.xlsm")
    reopened.run("Sheet1.ExportMarker")
    assert reopened.workbook.sheets_[1].Range("B1").vba_get("Value") == 0


def test_copy_preserves_unloaded_document_source_without_running_it(tmp_path: Path) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(MEASURED["document"], name="Sheet1", kind="document")
    source = app.save(tmp_path / "source.xlsm")
    unloaded = ExcelApplication.open(source, with_vba=False)
    unloaded.workbook.sheets_[0].Copy(After=unloaded.workbook.sheets_[0])
    assert unloaded.workbook.sheets_[1].vba_document is None
    unloaded.workbook.sheets_[1].Copy(After=unloaded.workbook.sheets_[1])
    copied = unloaded.save(tmp_path / "unloaded-copy.xlsm")
    reopened = ExcelApplication.open(copied)
    reopened.run("Sheet2.Mark", 23)
    reopened.run("Sheet2.ExportMarker")
    assert reopened.workbook.sheets_[1].Range("B1").vba_get("Value") == 23
    reopened.run("Sheet3.Mark", 31)
    reopened.run("Sheet3.ExportMarker")
    assert reopened.workbook.sheets_[2].Range("B1").vba_get("Value") == 31
    unloaded.add_module('Public Sub ExportMarker()\nRange("B1").Value = 71\nEnd Sub', name="Sheet2", kind="document")
    edited = unloaded.save(tmp_path / "edited.xlsm")
    edited_app = ExcelApplication.open(edited)
    edited_app.run("Sheet2.ExportMarker")
    assert edited_app.workbook.sheets_[1].Range("B1").vba_get("Value") == 71


@pytest.mark.skipif(os.environ.get("RUN_LIVE_EXCEL") != "1", reason="requires isolated real Excel")
def test_native_excel_executes_copied_document_and_deleted_module_is_gone(tmp_path: Path) -> None:
    harness = importlib.import_module("pyvbaharness")
    app = ExcelApplication()
    book = app.add_workbook()
    app.add_module(MEASURED["document"], name="Sheet1", kind="document")
    original = book.sheets_[0]
    original.Copy(After=original)
    copied_path = app.save(tmp_path / "copied.xlsm")
    book.sheets_[1].Delete()
    deleted_path = app.save(tmp_path / "deleted.xlsm")
    source = ('Function Probe() As String\nDim book As Workbook, part As Object, found As Boolean\n'
              'Application.EnableEvents = False\nSet book = Workbooks.Open("' + str(copied_path).replace('"', '""') + '")\n'
              'Application.Run "\'copied.xlsm\'!Sheet2.Mark", 19\nApplication.Run "\'copied.xlsm\'!Sheet2.ExportMarker"\n'
              'Probe = CStr(book.Worksheets(2).Range("B1").Value)\nbook.Close False\n'
              'Set book = Workbooks.Open("' + str(deleted_path).replace('"', '""') + '")\n'
              'For Each part In book.VBProject.VBComponents\nIf part.Name = "Sheet2" Then found = True\nNext\n'
              'Probe = Probe & ":" & CStr(found)\nbook.Close False\nEnd Function')
    with harness.ExcelSession(harness.HarnessConfig(lock_wait_s=30.0)) as excel:
        excel.new_document()
        result = excel.run_vba(source, "Probe", timeout=30.0)
        assert result.ok, result
        assert not result.dialogs
        assert result.value == "19:False"
