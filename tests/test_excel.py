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
