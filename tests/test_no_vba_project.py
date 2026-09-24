"""A file saved before its first macro has no VBA project, which is its normal shape.

tests/fixtures/no_vba/ holds what scripts/measure_no_vba.py had Excel,
Word and PowerPoint save before a line of VBA was written, macro-enabled
and binary: no vbaProject.bin in the zip files, and no storage at all at
the root of the binary ones (no_vba.json). Each opens. Listing reads
answer empty, and a read or write that needs the project raises
NoVBAProjectError, naming the file and the application its first macro
has to be written in. A save writes the file out as it was read. A
damaged project is still an error of its own, never taken for none.
"""

from __future__ import annotations

import io
import json
import shutil
import zipfile
from collections.abc import Callable
from pathlib import Path
from typing import Any

import pytest

from pyopenvba import (
    ExcelFile,
    NoVBAProjectError,
    PowerPointFile,
    PyOpenVBAError,
    VBAProjectError,
    WordFile,
    pull,
    push,
)
from pyopenvba.cfb import CFB
from pyopenvba.vba import SignatureInfo

FIXTURES = Path(__file__).parent / "fixtures" / "no_vba"
RECORD: dict[str, dict[str, Any]] = json.loads((FIXTURES / "no_vba.json").read_text(encoding="utf-8"))
OPENERS: dict[str, tuple[type[ExcelFile] | type[WordFile] | type[PowerPointFile], str]] = {
    ".xlsm": (ExcelFile, "Excel"), ".xls": (ExcelFile, "Excel"), ".docm": (WordFile, "Word"),
    ".doc": (WordFile, "Word"), ".pptm": (PowerPointFile, "PowerPoint"), ".ppt": (PowerPointFile, "PowerPoint")}
Host = ExcelFile | WordFile | PowerPointFile


@pytest.mark.parametrize("name", list(RECORD))
def test_a_file_with_no_project_opens_and_lists_nothing(name: str, tmp_path: Path) -> None:
    opener, _ = OPENERS[Path(name).suffix]
    with opener(FIXTURES / name) as office_file:
        assert not office_file.has_vba_project()
        assert office_file.module_names() == []
        assert office_file.vba_modules() == {}
        assert office_file.forms() == []
        assert office_file.references() == []
        assert office_file.validate() == []
        assert office_file.vba_signature() == SignatureInfo()
        assert office_file.pull_modules(tmp_path / "pulled") == []
    assert list((tmp_path / "pulled").iterdir()) == []


@pytest.mark.parametrize("name", list(RECORD))
def test_what_needs_the_project_refuses_naming_the_file(name: str, tmp_path: Path) -> None:
    opener, application = OPENERS[Path(name).suffix]
    folder = tmp_path / "modules"
    folder.mkdir()
    (folder / "Module1.bas").write_text('Attribute VB_Name = "Module1"\nSub A()\nEnd Sub\n', encoding="utf-8")
    calls: dict[str, Callable[[Host], object]] = {
        "vba_project": lambda host: host.vba_project(),
        "vba_project_bytes": lambda host: host.vba_project_bytes(),
        "get_module": lambda host: host.get_module("Module1"),
        "set_module": lambda host: host.set_module("Module1", "Sub A()\r\nEnd Sub\r\n"),
        "push_modules": lambda host: host.push_modules(folder),
        "add_form": lambda host: host.add_form("UserForm1"),
        "add_reference": lambda host: host.add_reference("Access"),
        "remove_reference": lambda host: host.remove_reference("Access"),
    }
    with opener(FIXTURES / name) as office_file:
        for call in calls.values():
            with pytest.raises(NoVBAProjectError, match=application) as raised:
                call(office_file)
            assert repr(name) in str(raised.value)
            assert isinstance(raised.value, VBAProjectError)


@pytest.mark.parametrize("name", list(RECORD))
def test_a_save_writes_the_file_out_as_it_was_read(name: str, tmp_path: Path) -> None:
    opener, _ = OPENERS[Path(name).suffix]
    copy, out = tmp_path / name, tmp_path / f"out{Path(name).suffix}"
    shutil.copyfile(FIXTURES / name, copy)
    with opener(copy) as office_file:
        office_file.save(out)
        office_file.save()
    assert out.read_bytes() == (FIXTURES / name).read_bytes()
    assert copy.read_bytes() == (FIXTURES / name).read_bytes()


def test_pull_exports_nothing_and_push_refuses(tmp_path: Path) -> None:
    workbook = tmp_path / "workbook.xlsm"
    shutil.copyfile(FIXTURES / "workbook.xlsm", workbook)
    assert pull(workbook, tmp_path / "pulled") == []
    source = tmp_path / "source"
    source.mkdir()
    (source / "Module1.bas").write_text('Attribute VB_Name = "Module1"\nSub A()\nEnd Sub\n', encoding="utf-8")
    with pytest.raises(NoVBAProjectError):
        push(source, workbook)
    assert workbook.read_bytes() == (FIXTURES / "workbook.xlsm").read_bytes()


def test_a_binary_file_whose_project_is_damaged_is_not_taken_for_one_with_none(tmp_path: Path) -> None:
    cfb = CFB.from_bytes((Path(__file__).parent / "live_excel_testing" / "xls_test.xls").read_bytes())
    cfb.remove_stream_in_storage("VBA", "dir")
    damaged = tmp_path / "damaged.xls"
    damaged.write_bytes(cfb.to_bytes())
    with ExcelFile(damaged) as workbook:
        assert workbook.has_vba_project()
        with pytest.raises(VBAProjectError) as raised:
            workbook.module_names()
    assert not isinstance(raised.value, NoVBAProjectError)


def test_a_zip_whose_project_part_is_damaged_is_not_taken_for_one_with_none(tmp_path: Path) -> None:
    buffer = io.BytesIO()
    with zipfile.ZipFile(FIXTURES / "workbook.xlsm") as original, zipfile.ZipFile(buffer, "w") as copy:
        for info in original.infolist():
            copy.writestr(info, original.read(info.filename))
        copy.writestr("xl/vbaProject.bin", b"not a compound file")
    damaged = tmp_path / "damaged.xlsm"
    damaged.write_bytes(buffer.getvalue())
    with ExcelFile(damaged) as workbook:
        assert workbook.has_vba_project()
        with pytest.raises(PyOpenVBAError) as raised:
            workbook.module_names()
    assert not isinstance(raised.value, NoVBAProjectError)


def test_the_in_memory_excel_opens_a_workbook_with_no_project() -> None:
    from pyopenvba.apps.excel import ExcelApplication

    app = ExcelApplication.open(FIXTURES / "workbook.xlsm")
    assert app.load_vba(FIXTURES / "workbook.xlsm") == []
    assert app.evaluate("ActiveWorkbook.Worksheets.Count") == 1
