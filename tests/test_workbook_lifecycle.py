"""Cancellable save and close behavior measured against real Excel."""

from __future__ import annotations

import json
from pathlib import Path

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBARuntimeError

MEASURED = json.loads((Path(__file__).parent / "fixtures/workbook_lifecycle.json").read_text(encoding="utf-8"))


@pytest.mark.parametrize("probe", MEASURED["probes"], ids=lambda probe: probe["name"])
def test_workbook_lifecycle(probe: dict[str, object], tmp_path: Path) -> None:
    app = ExcelApplication()
    book = app.add_workbook()
    app.add_module(MEASURED["log"], name="EventLog")
    app.add_module(MEASURED["sink"], name="LifecycleSink", kind="class")
    app.add_module(MEASURED["book"], name=book.code_name, kind="document")
    target = tmp_path / "output.xlsm"
    source = str(probe["source"]).replace("{OUTPUT}", str(target))
    app.add_module(source, name="ProbeModule")
    assert app.run("Probe") == probe["result"]
    outputs = [path.name[len(target.name):] for path in tmp_path.glob(target.name + "*") if path.is_file()]
    assert outputs == probe["output_suffixes"]


def test_macro_free_save_removes_project_and_signature_parts(tmp_path: Path) -> None:
    app = ExcelApplication.open(Path(__file__).parent / "fixtures/shapes/excel_shapes.xlsm", with_vba=False)
    package = app.workbook.package
    drawing = package.read("xl/drawings/drawing1.xml")
    package.write("xl/vbaProjectSignature.bin", b"signature")
    package.write("xl/_rels/vbaProject.bin.rels", b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignature" Target="vbaProjectSignature.bin"/></Relationships>')
    content = package.read("[Content_Types].xml").decode("utf-8")
    package.write("[Content_Types].xml", content.replace("</Types>", '<Override PartName="/xl/vbaProjectSignature.bin" ContentType="application/vnd.ms-office.vbaProjectSignature"/></Types>').encode("utf-8"))
    target = tmp_path / "converted.xlsx"
    app.workbook.SaveAs(str(target), 51)
    assert not package.has("xl/vbaProject.bin")
    assert not package.has("xl/_rels/vbaProject.bin.rels")
    assert not package.has("xl/vbaProjectSignature.bin")
    assert "vbaProjectSignature" not in package.read("[Content_Types].xml").decode("utf-8")
    assert package.read("xl/drawings/drawing1.xml") == drawing
    reopened = ExcelApplication.open(target, with_vba=False)
    assert reopened.workbook.FileFormat() == 51
    reopened.workbook.SaveAs(str(tmp_path / "back.xlsm"), 52)
    assert ExcelApplication.open(tmp_path / "back.xlsm", with_vba=False).workbook.FileFormat() == 52


def test_save_as_handles_reordered_and_prefixed_content_types(tmp_path: Path) -> None:
    app = ExcelApplication()
    app.add_workbook()
    plain = tmp_path / "plain.xlsx"
    app.workbook.SaveAs(str(plain), 51)
    package = app.workbook.package
    content = package.read("[Content_Types].xml").decode("utf-8")
    content = content.replace('PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"', "ContentType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml' PartName='/xl/workbook.xml'")
    content = content.replace("<Override ", '<ct:Override xmlns:ct="http://schemas.openxmlformats.org/package/2006/content-types" ')
    package.write("[Content_Types].xml", content.encode("utf-8"))
    target = tmp_path / "converted.xlsm"
    app.workbook.SaveAs(str(target), 52)
    reopened = ExcelApplication.open(target, with_vba=False)
    assert reopened.workbook.FileFormat() == 52


def test_python_save_respects_cancelled_save(tmp_path: Path) -> None:
    app = ExcelApplication()
    book = app.add_workbook()
    app.add_module("Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)\nCancel = True\nEnd Sub", name=book.code_name, kind="document")
    path = tmp_path / "cancelled.xlsx"
    with pytest.raises(VBARuntimeError, match="cancelled") as caught:
        app.save(path)
    assert caught.value.number == 1004
    assert not path.exists()
    assert book.path == ""
