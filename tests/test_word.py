"""Tests for the WordFile public interface."""

from __future__ import annotations

import io
import zipfile
from pathlib import Path

import pytest

from pyopenvba.exceptions import UnsupportedFormatError, VBAProjectError
from pyopenvba.word import WordFile

_VBA_ENTRY = "word/vbaProject.bin"

_LIVE_DOCM = Path(__file__).parent / "live_word_testing" / "Doc1.docm"


def _make_empty_zip_docm(tmp_path: Path, include_vba: bool = True) -> Path:
    """Build a fake .docm (ZIP) with an optional word/vbaProject.bin entry."""
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as zf:
        zf.writestr("[Content_Types].xml", "<Types/>")
        if include_vba:
            zf.writestr(_VBA_ENTRY, b"\x00" * 8)
    path = tmp_path / "doc.docm"
    path.write_bytes(buf.getvalue())
    return path


# ---------------------------------------------------------------------------
# Open / format detection
# ---------------------------------------------------------------------------

class TestWordFileOpen:
    def test_unsupported_extension_raises(self, tmp_path: Path) -> None:
        p = tmp_path / "doc.txt"
        p.write_bytes(b"hello")
        with pytest.raises(UnsupportedFormatError, match=r"\.txt"):
            WordFile(p)

    def test_docm_without_vba_entry_raises(self, tmp_path: Path) -> None:
        path = _make_empty_zip_docm(tmp_path, include_vba=False)
        with pytest.raises(VBAProjectError, match=r"vbaProject\.bin"):
            WordFile(path)

    def test_context_manager_docm(self, tmp_path: Path) -> None:
        path = _make_empty_zip_docm(tmp_path)
        with WordFile(path) as doc:
            assert doc is not None

    def test_save_invalid_vba_bin_raises(self, tmp_path: Path) -> None:
        from pyopenvba.exceptions import CFBError
        path = _make_empty_zip_docm(tmp_path)
        with WordFile(path) as doc, pytest.raises(CFBError):
            doc.save(tmp_path / "out.docm")

    def test_dotm_extension_accepted(self, tmp_path: Path) -> None:
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w") as zf:
            zf.writestr("[Content_Types].xml", "<Types/>")
            zf.writestr(_VBA_ENTRY, b"\x00" * 8)
        path = tmp_path / "template.dotm"
        path.write_bytes(buf.getvalue())
        with WordFile(path) as doc:
            assert doc is not None

    def test_doc_extension_raises_cfb_on_bad_data(self, tmp_path: Path) -> None:
        # A real .doc file must be a valid CFB; an 8-zero-byte file is not.
        # _open_cfb_direct() parses immediately, so CFBError surfaces in __init__.
        from pyopenvba.exceptions import CFBError
        p = tmp_path / "doc.doc"
        p.write_bytes(b"\x00" * 8)
        with pytest.raises(CFBError):
            WordFile(p)


# ---------------------------------------------------------------------------
# create_new
# ---------------------------------------------------------------------------

class TestWordFileCreateNew:
    def test_create_new_writes_file(self, tmp_path: Path) -> None:
        target = tmp_path / "new_doc.docm"
        doc = WordFile.create_new(target)
        try:
            assert target.exists()
            assert target.stat().st_size > 0
        finally:
            doc.close()

    def test_create_new_has_expected_modules(self, tmp_path: Path) -> None:
        target = tmp_path / "new_doc.docm"
        with WordFile.create_new(target) as doc:
            names = {m.name for m in doc.vba_project().modules}
        assert {"ThisDocument", "Module1"}.issubset(names)

    def test_create_new_module1_is_empty(self, tmp_path: Path) -> None:
        target = tmp_path / "new_doc.docm"
        with WordFile.create_new(target) as doc:
            src = doc.get_module("Module1")
        # Only the attribute header line; no body code.
        assert 'Attribute VB_Name = "Module1"' in src
        assert "Sub " not in src
        assert "Function " not in src

    def test_create_new_thisdocument_has_vb_base(self, tmp_path: Path) -> None:
        target = tmp_path / "new_doc.docm"
        with WordFile.create_new(target) as doc:
            src = doc.get_module("ThisDocument")
        assert "VB_Base" in src

    def test_create_new_overwrites_existing(self, tmp_path: Path) -> None:
        target = tmp_path / "new_doc.docm"
        target.write_bytes(b"old content")
        with WordFile.create_new(target) as doc:
            assert doc.vba_project() is not None

    def test_create_new_roundtrip_edit(self, tmp_path: Path) -> None:
        target = tmp_path / "new_doc.docm"
        out    = tmp_path / "edited.docm"
        marker = "'create_new roundtrip\r\n"
        with WordFile.create_new(target) as doc:
            doc.set_module("Module1", marker)
            doc.save(out)
        with WordFile(out) as doc2:
            assert marker in doc2.get_module("Module1")


# ---------------------------------------------------------------------------
# Live fixture round-trip (skipped when fixture is absent)
# ---------------------------------------------------------------------------

@pytest.mark.skipif(not _LIVE_DOCM.exists(), reason="live fixture not present")
class TestWordFileLive:
    def test_vba_modules_returns_dict(self) -> None:
        with WordFile(_LIVE_DOCM) as doc:
            modules = doc.vba_modules()
        assert isinstance(modules, dict)
        assert len(modules) > 0

    def test_module_names_match_vba_modules_keys(self) -> None:
        with WordFile(_LIVE_DOCM) as doc:
            names = doc.module_names()
            modules = doc.vba_modules()
        assert set(names) == set(modules.keys())

    def test_get_module_returns_source(self) -> None:
        with WordFile(_LIVE_DOCM) as doc:
            names = doc.module_names()
            src = doc.get_module(names[0])
        assert isinstance(src, str)

    def test_get_module_missing_raises(self) -> None:
        with WordFile(_LIVE_DOCM) as doc, pytest.raises(KeyError):
            doc.get_module("__nonexistent__")

    def test_set_module_marks_dirty(self) -> None:
        with WordFile(_LIVE_DOCM) as doc:
            proj = doc.vba_project()
            name = proj.modules[0].name
            doc.set_module(name, "Sub Hello()\nEnd Sub\n")
            assert any(m.dirty for m in proj.modules)

    def test_pull_modules_writes_files(self, tmp_path: Path) -> None:
        with WordFile(_LIVE_DOCM) as doc:
            written = doc.pull_modules(tmp_path)
        assert len(written) > 0
        for p in written:
            assert p.exists()
            assert p.suffix in {".bas", ".cls"}

    def test_roundtrip_source_edit(self, tmp_path: Path) -> None:
        """Edit a module, save to a temp file, reload and verify the change."""
        out = tmp_path / "edited.docm"
        marker = "'pyOpenVBA roundtrip marker\r\n"
        with WordFile(_LIVE_DOCM) as doc:
            names = doc.module_names()
            name = names[0]
            original = doc.get_module(name)
            doc.set_module(name, original + marker)
            doc.save(out)

        with WordFile(out) as doc2:
            result = doc2.get_module(name)
        assert marker in result

    def test_pull_then_push_roundtrip(self, tmp_path: Path) -> None:
        """pull -> mutate on disk -> push -> reload and verify."""
        src_dir = tmp_path / "src"
        out = tmp_path / "pushed.docm"

        with WordFile(_LIVE_DOCM) as doc:
            doc.pull_modules(src_dir)

        # Append a comment to the first .bas or .cls file found.
        bas_files = sorted(src_dir.glob("*.bas")) + sorted(src_dir.glob("*.cls"))
        assert bas_files, "No .bas/.cls files pulled"
        target_file = bas_files[0]
        module_name = target_file.stem
        marker = "'push roundtrip marker\r\n"
        target_file.write_bytes(target_file.read_bytes() + marker.encode())

        import shutil
        shutil.copy(_LIVE_DOCM, out)
        with WordFile(out) as doc:
            doc.push_modules(src_dir)
            doc.save()

        with WordFile(out) as doc:
            result = doc.get_module(module_name)
        assert marker in result

    def test_validate_returns_list(self) -> None:
        with WordFile(_LIVE_DOCM) as doc:
            issues = doc.validate()
        assert isinstance(issues, list)

    def test_vba_project_bytes_returns_bytes(self) -> None:
        with WordFile(_LIVE_DOCM) as doc:
            raw = doc.vba_project_bytes()
        assert isinstance(raw, bytes)
        assert len(raw) > 0


# ---------------------------------------------------------------------------
# All module kinds — combined inject + roundtrip
# ---------------------------------------------------------------------------

class TestWordAllModuleTypesLive:
    """Inject into every module kind that exists in a Word VBA project and
    verify the source survives a full save → reload cycle."""

    def test_standard_module_inject_roundtrip(self, tmp_path: Path) -> None:
        target = tmp_path / "word_std.docm"
        out    = tmp_path / "word_std_out.docm"
        body   = "Function DoubleIt(n As Long) As Long\r\n    DoubleIt = n * 2\r\nEnd Function\r\n"
        with WordFile.create_new(target) as doc:
            doc.set_module("Module1", body)
            doc.save(out)
        with WordFile(out) as doc2:
            src = doc2.get_module("Module1")
        assert "DoubleIt" in src

    def test_document_module_body_only_preserves_vb_base(self, tmp_path: Path) -> None:
        """ThisDocument is kind=other with VB_Base; body-only set must keep the header."""
        target = tmp_path / "word_td.docm"
        out    = tmp_path / "word_td_out.docm"
        body   = "Sub Document_Open()\r\n    MsgBox \"Opened!\"\r\nEnd Sub\r\n"
        with WordFile.create_new(target) as doc:
            assert "VB_Base" in doc.get_module("ThisDocument")
            doc.set_module("ThisDocument", body)
            doc.save(out)
        with WordFile(out) as doc2:
            src = doc2.get_module("ThisDocument")
        assert "VB_Base" in src, "VB_Base lost after body-only set_module on ThisDocument"
        assert "Document_Open" in src

    def test_all_module_types_combined_roundtrip(self, tmp_path: Path) -> None:
        """Inject into Module1 (standard) and ThisDocument (other) in one session."""
        target   = tmp_path / "word_all.docm"
        out      = tmp_path / "word_all_out.docm"
        std_body = "Function Cube(n As Long) As Long\r\n    Cube = n * n * n\r\nEnd Function\r\n"
        doc_body = "Sub Document_Close()\r\n    MsgBox \"Closing!\"\r\nEnd Sub\r\n"
        with WordFile.create_new(target) as doc:
            doc.set_module("Module1",     std_body)
            doc.set_module("ThisDocument", doc_body)
            doc.save(out)
        with WordFile(out) as doc2:
            mod1     = doc2.get_module("Module1")
            this_doc = doc2.get_module("ThisDocument")
        assert "Cube" in mod1
        assert "Document_Close" in this_doc
        assert "VB_Base" in this_doc

    @pytest.mark.skipif(not _LIVE_DOCM.exists(), reason="live fixture not present")
    def test_live_fixture_all_module_types_roundtrip(self, tmp_path: Path) -> None:
        """Same combined inject from the real Doc1.docm fixture."""
        out      = tmp_path / "word_live_all.docm"
        std_body = "Function TripleIt(n As Long) As Long\r\n    TripleIt = n * 3\r\nEnd Function\r\n"
        doc_body = "Sub Document_New()\r\n    MsgBox \"New document!\"\r\nEnd Sub\r\n"
        with WordFile(_LIVE_DOCM) as doc:
            doc.set_module("Module1",     std_body)
            doc.set_module("ThisDocument", doc_body)
            doc.save(out)
        with WordFile(out) as doc2:
            mod1     = doc2.get_module("Module1")
            this_doc = doc2.get_module("ThisDocument")
        assert "TripleIt"      in mod1
        assert "Document_New"  in this_doc
        assert "VB_Base"       in this_doc


# ---------------------------------------------------------------------------
# pull_word / push_word convenience helpers
# ---------------------------------------------------------------------------

@pytest.mark.skipif(not _LIVE_DOCM.exists(), reason="live fixture not present")
class TestPullPushWordHelpers:
    def test_pull_word_helper(self, tmp_path: Path) -> None:
        from pyopenvba import pull_word
        written = pull_word(_LIVE_DOCM, tmp_path)
        assert len(written) > 0

    def test_push_word_helper(self, tmp_path: Path) -> None:
        import shutil

        from pyopenvba import pull_word, push_word

        src = tmp_path / "src"
        out = tmp_path / "out.docm"
        shutil.copy(_LIVE_DOCM, out)

        pull_word(_LIVE_DOCM, src)

        bas_files = sorted(src.glob("*.bas")) + sorted(src.glob("*.cls"))
        if bas_files:
            marker = "'helper push marker\r\n"
            bas_files[0].write_bytes(bas_files[0].read_bytes() + marker.encode())

        updated = push_word(src, out)
        assert isinstance(updated, list)


class TestWordClassSourceNormalization:
    """set_module on a class-kind target normalizes VBE export form
    (GitHub issue #1); WordFile carries its own copy of set_module."""

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
        target = tmp_path / "doc.docm"
        with WordFile.create_new(target) as doc:
            doc.vba_project().add_module(
                "Class1", "Private x As Long\r\n", kind=VBAModuleKind.other
            )
            doc.save()
        with WordFile(target) as doc:
            doc.set_module("Class1", export_form)
            source = doc.get_module("Class1")
        assert not source.startswith("VERSION")
        header, _ = split_attribute_header(source)
        assert f'Attribute VB_Base = "{CLASS_MODULE_CLSID}"' in header
