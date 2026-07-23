"""Tests for the PowerPointFile public interface."""

from __future__ import annotations

import io
import zipfile
from pathlib import Path

import pytest

from pyopenvba.exceptions import UnsupportedFormatError, VBAProjectError
from pyopenvba.powerpoint import PowerPointFile

_VBA_ENTRY = "ppt/vbaProject.bin"

_LIVE_PPTM = Path(__file__).parent / "live_powerpoint_testing" / "Presentation1.pptm"


def _make_empty_zip_pptm(tmp_path: Path, include_vba: bool = True) -> Path:
    """Build a fake .pptm (ZIP) with an optional ppt/vbaProject.bin entry."""
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as zf:
        zf.writestr("[Content_Types].xml", "<Types/>")
        if include_vba:
            zf.writestr(_VBA_ENTRY, b"\x00" * 8)
    path = tmp_path / "pres.pptm"
    path.write_bytes(buf.getvalue())
    return path


# ---------------------------------------------------------------------------
# Open / format detection
# ---------------------------------------------------------------------------

class TestPowerPointFileOpen:
    def test_unsupported_extension_raises(self, tmp_path: Path) -> None:
        p = tmp_path / "pres.txt"
        p.write_bytes(b"hello")
        with pytest.raises(UnsupportedFormatError, match=r"\.txt"):
            PowerPointFile(p)

    def test_pptm_without_vba_entry_raises(self, tmp_path: Path) -> None:
        path = _make_empty_zip_pptm(tmp_path, include_vba=False)
        with pytest.raises(VBAProjectError, match=r"vbaProject\.bin"):
            PowerPointFile(path)

    def test_context_manager_pptm(self, tmp_path: Path) -> None:
        path = _make_empty_zip_pptm(tmp_path)
        with PowerPointFile(path) as prs:
            assert prs is not None

    def test_save_invalid_vba_bin_raises(self, tmp_path: Path) -> None:
        from pyopenvba.exceptions import CFBError
        path = _make_empty_zip_pptm(tmp_path)
        with PowerPointFile(path) as prs, pytest.raises(CFBError):
            prs.save(tmp_path / "out.pptm")

    def test_potm_extension_accepted(self, tmp_path: Path) -> None:
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w") as zf:
            zf.writestr("[Content_Types].xml", "<Types/>")
            zf.writestr(_VBA_ENTRY, b"\x00" * 8)
        path = tmp_path / "template.potm"
        path.write_bytes(buf.getvalue())
        with PowerPointFile(path) as prs:
            assert prs is not None

    def test_ppt_extension_raises_cfb_on_bad_data(self, tmp_path: Path) -> None:
        from pyopenvba.exceptions import CFBError
        p = tmp_path / "pres.ppt"
        p.write_bytes(b"\x00" * 8)
        with pytest.raises(CFBError):
            PowerPointFile(p)


# ---------------------------------------------------------------------------
# create_new
# ---------------------------------------------------------------------------

class TestPowerPointFileCreateNew:
    def test_create_new_writes_file(self, tmp_path: Path) -> None:
        target = tmp_path / "new_pres.pptm"
        prs = PowerPointFile.create_new(target)
        try:
            assert target.exists()
            assert target.stat().st_size > 0
        finally:
            prs.close()

    def test_create_new_has_module1(self, tmp_path: Path) -> None:
        target = tmp_path / "new_pres.pptm"
        with PowerPointFile.create_new(target) as prs:
            names = {m.name for m in prs.vba_project().modules}
        assert "Module1" in names

    def test_create_new_module1_is_empty(self, tmp_path: Path) -> None:
        target = tmp_path / "new_pres.pptm"
        with PowerPointFile.create_new(target) as prs:
            src = prs.get_module("Module1")
        assert 'Attribute VB_Name = "Module1"' in src
        assert "Sub " not in src
        assert "Function " not in src

    def test_create_new_overwrites_existing(self, tmp_path: Path) -> None:
        target = tmp_path / "new_pres.pptm"
        target.write_bytes(b"old content")
        with PowerPointFile.create_new(target) as prs:
            assert prs.vba_project() is not None

    def test_create_new_roundtrip_edit(self, tmp_path: Path) -> None:
        target = tmp_path / "new_pres.pptm"
        out    = tmp_path / "edited.pptm"
        marker = "'create_new roundtrip\r\n"
        with PowerPointFile.create_new(target) as prs:
            prs.set_module("Module1", marker)
            prs.save(out)
        with PowerPointFile(out) as prs2:
            assert marker in prs2.get_module("Module1")


# ---------------------------------------------------------------------------
# Live fixture round-trip (skipped when fixture is absent)
# ---------------------------------------------------------------------------

@pytest.mark.skipif(not _LIVE_PPTM.exists(), reason="live fixture not present")
class TestPowerPointFileLive:
    def test_vba_modules_returns_dict(self) -> None:
        with PowerPointFile(_LIVE_PPTM) as prs:
            modules = prs.vba_modules()
        assert isinstance(modules, dict)
        assert len(modules) > 0

    def test_module_names_match_vba_modules_keys(self) -> None:
        with PowerPointFile(_LIVE_PPTM) as prs:
            names   = prs.module_names()
            modules = prs.vba_modules()
        assert set(names) == set(modules.keys())

    def test_get_module_returns_source(self) -> None:
        with PowerPointFile(_LIVE_PPTM) as prs:
            name = prs.module_names()[0]
            src  = prs.get_module(name)
        assert isinstance(src, str)

    def test_get_module_missing_raises(self) -> None:
        with PowerPointFile(_LIVE_PPTM) as prs, pytest.raises(KeyError):
            prs.get_module("__nonexistent__")

    def test_set_module_marks_dirty(self) -> None:
        with PowerPointFile(_LIVE_PPTM) as prs:
            proj = prs.vba_project()
            name = proj.modules[0].name
            prs.set_module(name, "Sub Hello()\nEnd Sub\n")
            assert any(m.dirty for m in proj.modules)

    def test_pull_modules_writes_files(self, tmp_path: Path) -> None:
        with PowerPointFile(_LIVE_PPTM) as prs:
            written = prs.pull_modules(tmp_path)
        assert len(written) > 0
        for p in written:
            assert p.exists()
            assert p.suffix in {".bas", ".cls"}

    def test_roundtrip_source_edit(self, tmp_path: Path) -> None:
        out    = tmp_path / "edited.pptm"
        marker = "'pyOpenVBA roundtrip marker\r\n"
        with PowerPointFile(_LIVE_PPTM) as prs:
            # Edit the first standard module.
            std = next(
                (m for m in prs.vba_project().modules
                 if m.kind.name == "standard"), None
            )
            assert std is not None, "No standard module in live fixture"
            name = std.name
            original = prs.get_module(name)
            prs.set_module(name, original + marker)
            prs.save(out)
        with PowerPointFile(out) as prs2:
            result = prs2.get_module(name)
        assert marker in result

    def test_pull_then_push_roundtrip(self, tmp_path: Path) -> None:
        import shutil
        src_dir = tmp_path / "src"
        out     = tmp_path / "pushed.pptm"

        with PowerPointFile(_LIVE_PPTM) as prs:
            prs.pull_modules(src_dir)

        bas_files = sorted(src_dir.glob("*.bas")) + sorted(src_dir.glob("*.cls"))
        assert bas_files, "No .bas/.cls files pulled"
        target_file = bas_files[0]
        module_name = target_file.stem
        marker = "'push roundtrip marker\r\n"
        target_file.write_bytes(target_file.read_bytes() + marker.encode())

        shutil.copy(_LIVE_PPTM, out)
        with PowerPointFile(out) as prs:
            prs.push_modules(src_dir)
            prs.save()

        with PowerPointFile(out) as prs:
            result = prs.get_module(module_name)
        assert marker in result

    def test_validate_returns_list(self) -> None:
        with PowerPointFile(_LIVE_PPTM) as prs:
            issues = prs.validate()
        assert isinstance(issues, list)

    def test_vba_project_bytes_returns_bytes(self) -> None:
        with PowerPointFile(_LIVE_PPTM) as prs:
            raw = prs.vba_project_bytes()
        assert isinstance(raw, bytes)
        assert len(raw) > 0


# ---------------------------------------------------------------------------
# All module kinds — combined inject + roundtrip
# ---------------------------------------------------------------------------

class TestPowerPointAllModuleTypesLive:
    """Inject into Module1 (the only module kind in the blank template) and
    verify source survives a full save -> reload cycle."""

    def test_create_new_module1_roundtrip(self, tmp_path: Path) -> None:
        target = tmp_path / "ppt_m1.pptm"
        out    = tmp_path / "ppt_m1_out.pptm"
        body   = "Sub ShowCount()\r\n    MsgBox ActivePresentation.Slides.Count\r\nEnd Sub\r\n"
        with PowerPointFile.create_new(target) as prs:
            prs.set_module("Module1", body)
            prs.save(out)
        with PowerPointFile(out) as prs2:
            assert "ShowCount" in prs2.get_module("Module1")

    def test_standard_module_kind_is_standard(self, tmp_path: Path) -> None:
        target = tmp_path / "ppt_kinds.pptm"
        with PowerPointFile.create_new(target) as prs:
            proj = prs.vba_project()
        from pyopenvba.vba import VBAModuleKind
        mod1 = next(m for m in proj.modules if m.name == "Module1")
        assert mod1.kind == VBAModuleKind.standard

    @pytest.mark.skipif(not _LIVE_PPTM.exists(), reason="live fixture not present")
    def test_live_fixture_module1_roundtrip(self, tmp_path: Path) -> None:
        out  = tmp_path / "ppt_live_m1.pptm"
        body = "Sub LiveTest()\r\n    MsgBox \"live\"\r\nEnd Sub\r\n"
        with PowerPointFile(_LIVE_PPTM) as prs:
            prs.set_module("Module1", body)
            prs.save(out)
        with PowerPointFile(out) as prs2:
            assert "LiveTest" in prs2.get_module("Module1")


# ---------------------------------------------------------------------------
# pull_ppt / push_ppt convenience helpers
# ---------------------------------------------------------------------------

@pytest.mark.skipif(not _LIVE_PPTM.exists(), reason="live fixture not present")
class TestPullPushPptHelpers:
    def test_pull_ppt_helper(self, tmp_path: Path) -> None:
        from pyopenvba import pull_ppt
        written = pull_ppt(_LIVE_PPTM, tmp_path)
        assert len(written) > 0

    def test_push_ppt_helper(self, tmp_path: Path) -> None:
        import shutil

        from pyopenvba import pull_ppt, push_ppt

        src = tmp_path / "src"
        out = tmp_path / "out.pptm"
        shutil.copy(_LIVE_PPTM, out)

        pull_ppt(_LIVE_PPTM, src)

        bas_files = sorted(src.glob("*.bas")) + sorted(src.glob("*.cls"))
        if bas_files:
            marker = "'helper push marker\r\n"
            bas_files[0].write_bytes(bas_files[0].read_bytes() + marker.encode())

        updated = push_ppt(src, out)
        assert isinstance(updated, list)


class TestPowerPointClassSourceNormalization:
    """set_module on a class-kind target normalizes VBE export form
    (GitHub issue #1); PowerPointFile carries its own copy of set_module."""

    def test_set_module_normalizes_export_form_class(self, tmp_path: Path) -> None:
        from pyopenvba.vba import (
            CLASS_MODULE_CLSID,
            VBAModuleKind,
            split_attribute_header,
        )

        export_form = (
            "VERSION 1.0 CLASS\r\n"
            "BEGIN\r\n"
            "  MultiUse = -1  'True\r\n"
            "END\r\n"
            'Attribute VB_Name = "Class1"\r\n'
            "Attribute VB_Exposed = False\r\n"
            "\r\n"
            "Private mMessage As String\r\n"
        )
        target = tmp_path / "prs.pptm"
        with PowerPointFile.create_new(target) as prs:
            prs.vba_project().add_module(
                "Class1", "Private x As Long\r\n", kind=VBAModuleKind.other
            )
            prs.save()
        with PowerPointFile(target) as prs:
            prs.set_module("Class1", export_form)
            source = prs.get_module("Class1")
        assert not source.startswith("VERSION")
        header, _ = split_attribute_header(source)
        assert f'Attribute VB_Base = "{CLASS_MODULE_CLSID}"' in header
