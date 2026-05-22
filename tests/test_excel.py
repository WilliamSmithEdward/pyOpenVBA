"""Tests for the ExcelFile public interface."""

from __future__ import annotations

import io
import zipfile
from pathlib import Path

import pytest

from pyopenvba.excel import ExcelFile
from pyopenvba.exceptions import UnsupportedFormatError, VBAProjectError

_VBA_ENTRY = "xl/vbaProject.bin"


def _make_empty_zip_xlsm(tmp_path: Path, include_vba: bool = True) -> Path:
    """
    Build a fake .xlsm (ZIP) file.
    If include_vba is False, the xl/vbaProject.bin entry is omitted.
    """
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as zf:
        zf.writestr("[Content_Types].xml", "<Types/>")
        if include_vba:
            # Minimal (non-parseable) placeholder for vbaProject.bin
            zf.writestr(_VBA_ENTRY, b"\x00" * 8)
    path = tmp_path / "book.xlsm"
    path.write_bytes(buf.getvalue())
    return path


class TestExcelFileOpen:
    def test_unsupported_extension_raises(self, tmp_path: Path) -> None:
        p = tmp_path / "book.csv"
        p.write_bytes(b"a,b,c")
        with pytest.raises(UnsupportedFormatError, match=".csv"):
            ExcelFile(p)

    def test_xlsm_without_vba_entry_raises(self, tmp_path: Path) -> None:
        path = _make_empty_zip_xlsm(tmp_path, include_vba=False)
        with pytest.raises(VBAProjectError, match="vbaProject.bin"):
            ExcelFile(path)

    def test_context_manager(self, tmp_path: Path) -> None:
        path = _make_empty_zip_xlsm(tmp_path)
        with ExcelFile(path) as wb:
            assert wb is not None

    def test_save_invalid_vba_bin_raises(self, tmp_path: Path) -> None:
        # The fake xlsm above only stores 8 zero bytes as vbaProject.bin,
        # so save() must fail at CFB parsing, not silently emit garbage.
        from pyopenvba.exceptions import CFBError
        path = _make_empty_zip_xlsm(tmp_path)
        with ExcelFile(path) as wb:
            with pytest.raises(CFBError):
                wb.save(tmp_path / "out.xlsm")


class TestExcelFileCreateNew:
    def test_create_new_writes_file(self, tmp_path: Path) -> None:
        target = tmp_path / "new_book.xlsm"
        wb = ExcelFile.create_new(target)
        try:
            assert target.exists()
            assert target.stat().st_size > 0
        finally:
            wb.close()

    def test_create_new_has_expected_modules(self, tmp_path: Path) -> None:
        target = tmp_path / "new_book.xlsm"
        with ExcelFile.create_new(target) as wb:
            proj = wb.vba_project()
            names = {m.name for m in proj.modules}
        assert {"ThisWorkbook", "Sheet1", "Module1"}.issubset(names)

    def test_create_new_module1_is_empty(self, tmp_path: Path) -> None:
        target = tmp_path / "new_book.xlsm"
        with ExcelFile.create_new(target) as wb:
            proj = wb.vba_project()
            m = proj.get_module("Module1")
        assert m.source == 'Attribute VB_Name = "Module1"\r\n'

    def test_create_new_round_trip_with_user_code(self, tmp_path: Path) -> None:
        target = tmp_path / "new_book.xlsm"
        with ExcelFile.create_new(target) as wb:
            proj = wb.vba_project()
            m = proj.get_module("Module1")
            m.source = (
                'Attribute VB_Name = "Module1"\r\n'
                "Public Sub Hello()\r\n"
                "    MsgBox \"Hi\"\r\n"
                "End Sub\r\n"
            )
            m.dirty = True
            wb.save()
        with ExcelFile(target) as wb2:
            proj2 = wb2.vba_project()
            m2 = proj2.get_module("Module1")
            assert "Public Sub Hello()" in m2.source
            assert "MsgBox" in m2.source

    def test_create_new_overwrites_existing(self, tmp_path: Path) -> None:
        target = tmp_path / "new_book.xlsm"
        target.write_bytes(b"junk")
        with ExcelFile.create_new(target) as wb:
            proj = wb.vba_project()
            assert any(m.name == "ThisWorkbook" for m in proj.modules)

    def test_create_new_creates_parent_dirs(self, tmp_path: Path) -> None:
        target = tmp_path / "nested" / "dir" / "new_book.xlsm"
        with ExcelFile.create_new(target) as wb:
            assert target.exists()
            assert wb.vba_project() is not None
