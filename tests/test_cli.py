"""Tests for the ``python -m pyopenvba`` command-line interface."""

from __future__ import annotations

from pathlib import Path

import pytest

from pyopenvba import ExcelFile, PowerPointFile, WordFile
from pyopenvba.__main__ import main

_ACCDB = (
    Path(__file__).parent
    / "live_access_test"
    / "New Microsoft Access Database.accdb"
)


class TestLs:
    def test_ls_lists_modules_of_xlsm(
        self, tmp_path: Path, capsys: pytest.CaptureFixture[str]
    ) -> None:
        target = tmp_path / "book.xlsm"
        with ExcelFile.create_new(target) as wb:
            wb.save()
        assert main(["ls", str(target)]) == 0
        out = capsys.readouterr().out
        assert "Module1" in out
        assert "ThisWorkbook" in out

    def test_ls_routes_word_by_suffix(
        self, tmp_path: Path, capsys: pytest.CaptureFixture[str]
    ) -> None:
        target = tmp_path / "doc.docm"
        with WordFile.create_new(target) as doc:
            doc.save()
        assert main(["ls", str(target)]) == 0
        assert "ThisDocument" in capsys.readouterr().out

    def test_unsupported_suffix_fails_with_exit_2(
        self, tmp_path: Path, capsys: pytest.CaptureFixture[str]
    ) -> None:
        assert main(["ls", str(tmp_path / "notes.txt")]) == 2
        assert "unsupported file type" in capsys.readouterr().err

    def test_xltm_is_not_advertised(
        self, tmp_path: Path, capsys: pytest.CaptureFixture[str]
    ) -> None:
        # .xltm was previously listed for disasm but rejected by ExcelFile;
        # the suffix map now only carries extensions the facades accept.
        assert main(["disasm", str(tmp_path / "t.xltm")]) == 2
        assert "unsupported file type" in capsys.readouterr().err


class TestPullPush:
    def test_pull_edit_push_round_trip_word(
        self, tmp_path: Path, capsys: pytest.CaptureFixture[str]
    ) -> None:
        target = tmp_path / "doc.docm"
        with WordFile.create_new(target) as doc:
            doc.set_module(
                "Module1", 'Sub Marker()\r\n    MsgBox "before"\r\nEnd Sub\r\n'
            )
            doc.save()

        vba_dir = tmp_path / "vba"
        assert main(["pull", str(target), str(vba_dir)]) == 0
        pulled = capsys.readouterr().out.strip().splitlines()
        assert any(line.endswith("Module1.bas") for line in pulled)

        module_file = vba_dir / "Module1.bas"
        module_file.write_bytes(
            module_file.read_bytes().replace(b'"before"', b'"after"')
        )
        assert main(["push", str(vba_dir), str(target)]) == 0
        assert "Module1" in capsys.readouterr().out

        with WordFile(target) as doc:
            assert '"after"' in doc.get_module("Module1")

    def test_pull_routes_powerpoint_by_suffix(
        self, tmp_path: Path, capsys: pytest.CaptureFixture[str]
    ) -> None:
        target = tmp_path / "prs.pptm"
        with PowerPointFile.create_new(target) as prs:
            prs.save()
        vba_dir = tmp_path / "vba"
        assert main(["pull", str(target), str(vba_dir)]) == 0
        assert (vba_dir / "Module1.bas").exists()
        capsys.readouterr()


@pytest.mark.skipif(not _ACCDB.exists(), reason="live .accdb fixture not present")
class TestAccessPull:
    def test_access_pull_classifies_class_modules(
        self, tmp_path: Path, capsys: pytest.CaptureFixture[str]
    ) -> None:
        dest = tmp_path / "access_vba"
        assert main(["access-pull", str(_ACCDB), str(dest)]) == 0
        printed = capsys.readouterr().out.strip().splitlines()
        written = sorted(p.name for p in dest.iterdir())
        assert written, "expected at least one exported module"
        assert len(printed) == len(written)
        # Classification comes from the dir-stream catalog: standard
        # modules as .bas, class modules as .cls.
        assert all(name.endswith((".bas", ".cls")) for name in written)


class TestForms:
    _NESTED = Path(__file__).parent / "live_excel_testing" / "nested_form.xlsm"
    _LEGACY_PPT = (
        Path(__file__).parent / "live_powerpoint_testing" / "legacy_macros.ppt"
    )

    @pytest.mark.skipif(not _NESTED.exists(), reason="nested form fixture not present")
    def test_forms_prints_the_control_tree(
        self, capsys: pytest.CaptureFixture[str]
    ) -> None:
        assert main(["forms", "--mask", str(self._NESTED)]) == 0
        out = capsys.readouterr().out
        assert "FrmNested" in out
        assert "MSForms.MultiPage" in out
        # Nesting shows as indentation, so a child must be further in
        # than the container that holds it.
        indent = {
            ln.strip().split()[0]: len(ln) - len(ln.lstrip())
            for ln in out.splitlines()
            if ln.strip() and "MSForms." in ln
        }
        assert indent["OptOne"] > indent["GroupBox"]
        assert indent["PageOneCheck"] > indent["Page1"]

    @pytest.mark.skipif(not _NESTED.exists(), reason="nested form fixture not present")
    def test_forms_names_the_properties(
        self, capsys: pytest.CaptureFixture[str]
    ) -> None:
        assert main(["forms", str(self._NESTED)]) == 0
        out = capsys.readouterr().out
        assert "Caption = 'Close'" in out
        assert "Font.FontName = 'Tahoma'" in out

    @pytest.mark.skipif(not _NESTED.exists(), reason="nested form fixture not present")
    def test_forms_mask_flag_reports_the_raw_mask(
        self, capsys: pytest.CaptureFixture[str]
    ) -> None:
        assert main(["forms", "--mask", str(self._NESTED)]) == 0
        out = capsys.readouterr().out
        # MorphData controls carry the wider mask, and the printed width
        # is what tells a reader which bit index they are looking at.
        assert "set=0x0000000080000101" in out   # InnerText, 8 bytes
        assert "set=0x00000028" in out           # TopLabel, 4 bytes

    @pytest.mark.skipif(
        not _LEGACY_PPT.exists(), reason="legacy .ppt fixture not present"
    )
    def test_forms_says_so_when_a_project_has_none(
        self, capsys: pytest.CaptureFixture[str]
    ) -> None:
        assert main(["forms", str(self._LEGACY_PPT)]) == 0
        assert "no UserForms" in capsys.readouterr().out

    def test_forms_rejects_an_unsupported_suffix(
        self, tmp_path: Path, capsys: pytest.CaptureFixture[str]
    ) -> None:
        assert main(["forms", str(tmp_path / "notes.txt")]) == 2
        assert "unsupported file type" in capsys.readouterr().err
