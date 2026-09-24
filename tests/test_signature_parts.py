"""A VBA signature kept beside vbaProject.bin, as Excel, Word and PowerPoint keep it and drop it.

tests/fixtures/signature_parts/ is what scripts/measure_signature_parts.py
saw. Each application saved, untouched, a file carrying the three
signature parts of an add-in Office installs; the fixture is that save
with placeholder bytes in the parts, which are Microsoft's. Each then
saved the file again after one line of its module changed, and dropped
the signature: its [Content_Types].xml and its project's relationships
part are recorded. Here the library reads the signature from the first
save and, changing the same module, writes what the second wrote.
"""

from __future__ import annotations

import json
import warnings
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba import ExcelFile, PowerPointFile, WordFile
from pyopenvba._package_signature import without_signature

FIXTURES = Path(__file__).parent / "fixtures" / "signature_parts"
RECORD: dict[str, dict[str, Any]] = json.loads((FIXTURES / "signature_parts.json").read_text(encoding="utf-8"))
OPENERS: dict[str, type[ExcelFile] | type[WordFile] | type[PowerPointFile]] = {
    "excel": ExcelFile, "word": WordFile, "powerpoint": PowerPointFile}


def _saved_texts(path: Path, project: str) -> dict[str, Any]:
    """What the measurement recorded of a save: [Content_Types].xml, the relationships part, the signature parts."""
    with zipfile.ZipFile(path) as package:
        names = package.namelist()
        rels = f"{project}/_rels/vbaProject.bin.rels"
        return {"content_types": package.read("[Content_Types].xml").decode("utf-8"),
                "relationships": package.read(rels).decode("utf-8") if rels in names else None,
                "parts": [name for name in names if "vbaProjectSignature" in name]}


@pytest.mark.parametrize("host", list(RECORD))
def test_the_signature_is_read_from_the_parts_beside_the_project(host: str) -> None:
    record = RECORD[host]
    with OPENERS[host](FIXTURES / record["file"]) as office_file:
        found = office_file.vba_signature()
    assert found.present
    assert sorted(found.parts) == sorted(record["untouched_save"]["parts"])
    assert sorted(found.kinds) == ["agile", "legacy", "v3"]


@pytest.mark.parametrize("host", list(RECORD))
def test_an_edit_drops_the_signature_as_the_application_drops_it(host: str, tmp_path: Path) -> None:
    record = RECORD[host]
    out = tmp_path / record["file"]
    with OPENERS[host](FIXTURES / record["file"]) as office_file:
        office_file.set_module(record["module"], office_file.get_module(record["module"]) + "' edited\r\n")
        with pytest.warns(UserWarning, match="signature"):
            office_file.save(out)
    assert _saved_texts(out, record["project"]) == record["edited_save"]
    with OPENERS[host](out) as saved:
        assert not saved.vba_signature().present
        assert saved.get_module(record["module"]).endswith("' edited\r\n")


@pytest.mark.parametrize("host", list(RECORD))
def test_a_save_that_changes_nothing_keeps_the_signature(host: str, tmp_path: Path) -> None:
    record = RECORD[host]
    out = tmp_path / record["file"]
    with OPENERS[host](FIXTURES / record["file"]) as office_file, warnings.catch_warnings():
        warnings.simplefilter("error")
        office_file.save(out)
    assert _saved_texts(out, record["project"]) == record["untouched_save"]


def test_allow_invalidate_signature_drops_it_without_a_warning(tmp_path: Path) -> None:
    record = RECORD["word"]
    out = tmp_path / record["file"]
    with WordFile(FIXTURES / record["file"]) as document, warnings.catch_warnings():
        warnings.simplefilter("error")
        document.set_module(record["module"], document.get_module(record["module"]) + "' quiet\r\n")
        document.save(out, allow_invalidate_signature=True)
    assert _saved_texts(out, record["project"]) == record["edited_save"]


def test_an_edit_to_a_form_alone_drops_the_signature(tmp_path: Path) -> None:
    """A form's design is part of what the agile and V3 signatures cover."""
    record = RECORD["excel"]
    signed, out = tmp_path / "signed.xlsm", tmp_path / "out.xlsm"
    parts: list[str] = record["untouched_save"]["parts"]
    with zipfile.ZipFile(FIXTURES / record["file"]) as source:
        carried = {name: source.read(name) for name in [*parts, "xl/_rels/vbaProject.bin.rels"]}
    overrides = "".join(f'<Override PartName="/{name}" ContentType="application/vnd.ms-office.'
                        f'{name.removeprefix("xl/").removesuffix(".bin")}"/>' for name in parts)
    nested = Path(__file__).parent / "live_excel_testing" / "nested_form.xlsm"
    with zipfile.ZipFile(nested) as original, zipfile.ZipFile(signed, "w", zipfile.ZIP_DEFLATED) as copy:
        for info in original.infolist():
            data = original.read(info.filename)
            if info.filename == "[Content_Types].xml":
                data = data.replace(b"</Types>", overrides.encode("utf-8") + b"</Types>")
            copy.writestr(info, data)
        for name, data in carried.items():
            copy.writestr(name, data)
    with ExcelFile(signed) as workbook:
        assert workbook.vba_signature().present
        workbook.forms()[0].control("TopLabel").set_property("Caption", "Hi")
        with pytest.warns(UserWarning, match="signature"):
            workbook.save(out)
    with zipfile.ZipFile(out) as package:
        assert not [name for name in package.namelist() if "vbaProjectSignature" in name]
        assert "xl/_rels/vbaProject.bin.rels" not in package.namelist()
        assert b"vbaProjectSignature" not in package.read("[Content_Types].xml")
    with ExcelFile(out) as workbook:
        assert workbook.forms()[0].control("TopLabel").get("Caption") == "Hi"


def _package(parts: dict[str, str]) -> tuple[list[str], dict[str, bytes]]:
    data = {name: text.encode("utf-8") for name, text in parts.items()}
    return list(data), data


def test_other_relationships_stay_numbered_from_one_in_their_order() -> None:
    names, data = _package({
        "[Content_Types].xml": '<Types><Override PartName="/xl/sig.bin" ContentType="s"/>'
                               '<Override PartName="/xl/other.xml" ContentType="o"/></Types>',
        "xl/_rels/vbaProject.bin.rels":
            '<Relationships><Relationship Id="rId7" Type="urn:first" Target="first.xml"/>'
            '<Relationship Target="/xl/sig.bin" Id="rId2" '
            'Type="http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignature"/>'
            '<Relationship Id="rId5" Type="urn:second" Target="second.xml"/></Relationships>',
        "xl/sig.bin": "x"})
    edits = without_signature(names, data.__getitem__, "xl/vbaProject.bin")
    assert edits == {
        "xl/sig.bin": None,
        "xl/_rels/vbaProject.bin.rels":
            b'<Relationships><Relationship Id="rId1" Type="urn:first" Target="first.xml"/>'
            b'<Relationship Id="rId2" Type="urn:second" Target="second.xml"/></Relationships>',
        "[Content_Types].xml": b'<Types><Override PartName="/xl/other.xml" ContentType="o"/></Types>'}


def test_a_project_whose_relationships_name_no_signature_is_left_alone() -> None:
    names, data = _package({
        "[Content_Types].xml": "<Types/>",
        "word/_rels/vbaProject.bin.rels":
            '<Relationships><Relationship Id="rId1" '
            'Type="http://schemas.microsoft.com/office/2006/relationships/wordVbaData" Target="vbaData.xml"/>'
            "</Relationships>"})
    assert without_signature(names, data.__getitem__, "word/vbaProject.bin") == {}
    assert without_signature(["[Content_Types].xml"], data.__getitem__, "xl/vbaProject.bin") == {}
