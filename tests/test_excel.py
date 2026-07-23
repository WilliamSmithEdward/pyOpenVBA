"""Tests for the ExcelFile public interface."""

from __future__ import annotations

import io
import zipfile
from pathlib import Path

import pytest

from pyopenvba.excel import ExcelFile
from pyopenvba.exceptions import UnsupportedFormatError, VBAProjectError
from pyopenvba.vba import (
    CLASS_MODULE_CLSID,
    VBAModuleKind,
    split_attribute_header,
)

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
        with pytest.raises(UnsupportedFormatError, match=r"\.csv"):
            ExcelFile(p)

    def test_xlsm_without_vba_entry_raises(self, tmp_path: Path) -> None:
        path = _make_empty_zip_xlsm(tmp_path, include_vba=False)
        with pytest.raises(VBAProjectError, match=r"vbaProject\.bin"):
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
        with ExcelFile(path) as wb, pytest.raises(CFBError):
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


def test_vba_module_disassemble_decodes_pcode() -> None:
    import pyopenvba

    fixture = Path("tests/live_excel_testing/large_vba_module.xlsm")
    if not fixture.exists():
        import pytest
        pytest.skip("large_vba_module.xlsm not present")
    with pyopenvba.ExcelFile(fixture) as wb:
        project = wb.vba_project()
    target = next(
        m for m in project.modules if m.name == "Large_Module_"
    )
    disasm = target.disassemble()
    assert disasm.cafe_offset >= 0
    assert disasm.num_lines > 0
    listing = disasm.to_annotated_listing(target.source)
    assert "; >   0: Attribute VB_Name" in listing
    # Re-running with to_listing() should also succeed and not contain
    # source comments.
    plain = disasm.to_listing()
    assert "; DisassembledModule" in plain
    assert "; >" not in plain


def test_vba_module_disassemble_returns_empty_when_no_prefix() -> None:
    from pyopenvba.vba import VBAModule

    mod = VBAModule(name="Empty", stream_name="Empty", source="")
    disasm = mod.disassemble()
    assert disasm.cafe_offset == -1
    assert disasm.num_lines == 0
    assert disasm.lines == ()


# ---------------------------------------------------------------------------
# Class-source normalization at the facade entry points (GitHub issue #1)
# ---------------------------------------------------------------------------

_VB_BASE_LINE = f'Attribute VB_Base = "{CLASS_MODULE_CLSID}"'

_EXPORT_FORM_CLS = (
    "VERSION 1.0 CLASS\r\n"
    "BEGIN\r\n"
    "  MultiUse = -1  'True\r\n"
    "END\r\n"
    'Attribute VB_Name = "Class1"\r\n'
    "Attribute VB_GlobalNameSpace = False\r\n"
    "Attribute VB_Creatable = False\r\n"
    "Attribute VB_PredeclaredId = False\r\n"
    "Attribute VB_Exposed = False\r\n"
    "\r\n"
    "Private mMessage As String\r\n"
    "\r\n"
    "Public Function Greet() As String\r\n"
    '    Greet = "Class1.Greet: " & mMessage\r\n'
    "End Function\r\n"
)


def _new_workbook_with_class(path: Path) -> None:
    """Create an .xlsm containing a correctly-formed Class1 and save it."""
    with ExcelFile.create_new(path) as wb:
        project = wb.vba_project()
        project.add_module(
            "Class1", "Private mMessage As String\r\n", kind=VBAModuleKind.other
        )
        wb.save()


class TestClassSourceNormalization:
    def test_set_module_normalizes_export_form_class(self, tmp_path: Path) -> None:
        target = tmp_path / "book.xlsm"
        _new_workbook_with_class(target)
        with ExcelFile(target) as wb:
            wb.set_module("Class1", _EXPORT_FORM_CLS)
            wb.save()
        with ExcelFile(target) as wb:
            source = wb.get_module("Class1")
        assert not source.startswith("VERSION")
        header, _ = split_attribute_header(source)
        assert _VB_BASE_LINE in header
        assert source.endswith("End Function\r\n")

    def test_set_module_preserves_document_module_host_clsid(
        self, tmp_path: Path
    ) -> None:
        target = tmp_path / "book.xlsm"
        with ExcelFile.create_new(target) as wb:
            project = wb.vba_project()
            prior_header = project.get_module("ThisWorkbook").attribute_header
            prior_vb_base = next(
                line
                for line in prior_header.splitlines()
                if line.startswith("Attribute VB_Base")
            )
            # A document module's host CLSID differs from the universal
            # class CLSID; the assertion below relies on that.
            assert CLASS_MODULE_CLSID not in prior_vb_base
            # Full-source replacement whose header lacks VB_Base entirely.
            supplied = (
                'Attribute VB_Name = "ThisWorkbook"\r\n'
                "\r\n"
                "Private Sub Workbook_Open()\r\nEnd Sub\r\n"
            )
            wb.set_module("ThisWorkbook", supplied)
            new_header = project.get_module("ThisWorkbook").attribute_header
        assert prior_vb_base in new_header
        assert CLASS_MODULE_CLSID not in new_header

    def test_push_modules_normalizes_export_form_cls(self, tmp_path: Path) -> None:
        target = tmp_path / "book.xlsm"
        _new_workbook_with_class(target)
        src_dir = tmp_path / "vba"
        src_dir.mkdir()
        (src_dir / "Class1.cls").write_bytes(_EXPORT_FORM_CLS.encode("utf-8"))
        with ExcelFile(target) as wb:
            updated = wb.push_modules(src_dir)
            assert updated == ["Class1"]
            module = wb.vba_project().get_module("Class1")
            assert module.dirty
            assert not module.source.startswith("VERSION")
            assert _VB_BASE_LINE in module.attribute_header
            wb.save()
        # Pushing the same file again normalizes to the identical text and
        # leaves the module clean.
        with ExcelFile(target) as wb:
            wb.push_modules(src_dir)
            assert not wb.vba_project().get_module("Class1").dirty

    def test_pull_then_push_round_trip_stays_clean(self, tmp_path: Path) -> None:
        target = tmp_path / "book.xlsm"
        _new_workbook_with_class(target)
        pulled = tmp_path / "pulled"
        with ExcelFile(target) as wb:
            wb.pull_modules(pulled)
            wb.push_modules(pulled)
            assert not any(m.dirty for m in wb.vba_project().modules)
