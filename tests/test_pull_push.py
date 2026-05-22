"""Tests for the disk-based push/pull workflow."""

from __future__ import annotations

import shutil
import zipfile
from pathlib import Path

import pytest

from pyopenvba import ExcelFile, pull, push

LIVE_XLSM = Path(__file__).parent / "live_excel_testing" / "test_macro_workbook.xlsm"


@pytest.fixture(scope="module")
def live_xlsm() -> Path:
    if not LIVE_XLSM.exists():
        pytest.skip(f"live workbook not available at {LIVE_XLSM}")
    return LIVE_XLSM


# ---------------------------------------------------------------------------
# Basic shape
# ---------------------------------------------------------------------------

class TestPull:
    def test_pull_writes_one_file_per_module(
        self, live_xlsm: Path, tmp_path: Path
    ) -> None:
        out = tmp_path / "vba_src"
        written = pull(live_xlsm, out)
        with ExcelFile(live_xlsm) as wb:
            expected_names = {m.name for m in wb.vba_project().modules}
        stems = {p.stem for p in written}
        assert stems == expected_names

    def test_pull_uses_bas_and_cls_extensions(
        self, live_xlsm: Path, tmp_path: Path
    ) -> None:
        out = tmp_path / "vba_src"
        written = pull(live_xlsm, out)
        suffixes = {p.suffix for p in written}
        assert suffixes <= {".bas", ".cls"}

    def test_pull_creates_dest_dir(self, live_xlsm: Path, tmp_path: Path) -> None:
        out = tmp_path / "nested" / "dir" / "that" / "does" / "not" / "exist"
        assert not out.exists()
        pull(live_xlsm, out)
        assert out.is_dir()

    def test_pull_uses_crlf_line_endings(
        self, live_xlsm: Path, tmp_path: Path
    ) -> None:
        out = tmp_path / "vba_src"
        written = pull(live_xlsm, out)
        for p in written:
            raw = p.read_bytes()
            # If the file contains any newlines, they must be CRLF.
            if b"\n" in raw:
                # No bare LF (every LF must be preceded by CR).
                assert b"\r\n" in raw
                bare_lfs = sum(
                    1
                    for i, b in enumerate(raw)
                    if b == 0x0A and (i == 0 or raw[i - 1] != 0x0D)
                )
                assert bare_lfs == 0, f"{p} has bare LFs"

    def test_pull_overwrite_false_refuses_existing(
        self, live_xlsm: Path, tmp_path: Path
    ) -> None:
        out = tmp_path / "vba_src"
        pull(live_xlsm, out)
        with pytest.raises(FileExistsError):
            pull(live_xlsm, out, overwrite=False)


# ---------------------------------------------------------------------------
# Push round-trip
# ---------------------------------------------------------------------------

class TestPushRoundTrip:
    def test_pull_then_push_unchanged_preserves_module_sources(
        self, live_xlsm: Path, tmp_path: Path
    ) -> None:
        # 1. Pull from the live workbook.
        src_dir = tmp_path / "vba_src"
        pull(live_xlsm, src_dir)

        # 2. Push back into a copy (no edits).
        work = tmp_path / "work.xlsm"
        shutil.copy(live_xlsm, work)
        push(src_dir, work)

        # 3. Sources match the original after the round-trip.
        with ExcelFile(live_xlsm) as a, ExcelFile(work) as b:
            assert a.vba_modules() == b.vba_modules()

    def test_push_then_pull_propagates_edit(
        self, live_xlsm: Path, tmp_path: Path
    ) -> None:
        src_dir = tmp_path / "vba_src"
        pull(live_xlsm, src_dir)

        # Find any .bas/.cls file and append a comment line.
        target = next(p for p in src_dir.iterdir()
                      if p.suffix in {".bas", ".cls"})
        original = target.read_text(encoding="utf-8")
        edited = original + "'pyopenvba push/pull test sentinel\r\n"
        target.write_text(edited, encoding="utf-8", newline="")

        work = tmp_path / "edited.xlsm"
        shutil.copy(live_xlsm, work)
        updated = push(src_dir, work)
        assert target.stem in updated

        # Re-pull and confirm the edit survived.
        repulled = tmp_path / "repulled"
        pull(work, repulled)
        roundtrip = (repulled / target.name).read_text(encoding="utf-8")
        assert "pyopenvba push/pull test sentinel" in roundtrip

    def test_push_writes_to_out_without_modifying_source(
        self, live_xlsm: Path, tmp_path: Path
    ) -> None:
        src_dir = tmp_path / "vba_src"
        pull(live_xlsm, src_dir)
        before = live_xlsm.read_bytes()
        out = tmp_path / "out.xlsm"
        push(src_dir, live_xlsm, out=out)
        assert live_xlsm.read_bytes() == before
        assert out.exists()

    def test_push_preserves_non_vba_zip_entries(
        self, live_xlsm: Path, tmp_path: Path
    ) -> None:
        src_dir = tmp_path / "vba_src"
        pull(live_xlsm, src_dir)
        work = tmp_path / "work.xlsm"
        shutil.copy(live_xlsm, work)
        push(src_dir, work)

        with zipfile.ZipFile(live_xlsm) as before, zipfile.ZipFile(work) as after:
            before_names = set(before.namelist())
            after_names = set(after.namelist())
            assert before_names == after_names
            for name in before_names:
                if name == "xl/vbaProject.bin":
                    continue
                assert before.read(name) == after.read(name), name


# ---------------------------------------------------------------------------
# Unmatched-file behavior
# ---------------------------------------------------------------------------

class TestPushUnmatched:
    def test_unmatched_file_ignored_by_default(
        self, live_xlsm: Path, tmp_path: Path
    ) -> None:
        src_dir = tmp_path / "vba_src"
        pull(live_xlsm, src_dir)
        (src_dir / "DoesNotExist.bas").write_bytes(b"Sub Foo()\r\nEnd Sub\r\n")
        work = tmp_path / "work.xlsm"
        shutil.copy(live_xlsm, work)
        # Should not raise.
        push(src_dir, work)

    def test_unmatched_file_raises_in_strict_mode(
        self, live_xlsm: Path, tmp_path: Path
    ) -> None:
        src_dir = tmp_path / "vba_src"
        pull(live_xlsm, src_dir)
        (src_dir / "DoesNotExist.bas").write_bytes(b"Sub Foo()\r\nEnd Sub\r\n")
        work = tmp_path / "work.xlsm"
        shutil.copy(live_xlsm, work)
        with pytest.raises(KeyError):
            push(src_dir, work, strict=True)

    def test_non_source_files_ignored(
        self, live_xlsm: Path, tmp_path: Path
    ) -> None:
        src_dir = tmp_path / "vba_src"
        pull(live_xlsm, src_dir)
        (src_dir / "README.txt").write_text("not VBA\n", encoding="utf-8")
        (src_dir / "notes.md").write_text("notes\n", encoding="utf-8")
        work = tmp_path / "work.xlsm"
        shutil.copy(live_xlsm, work)
        # Should not raise even in strict mode (non-source ignored).
        push(src_dir, work, strict=True)

    def test_push_src_dir_not_a_directory(
        self, live_xlsm: Path, tmp_path: Path
    ) -> None:
        bogus = tmp_path / "not_a_dir.txt"
        bogus.write_text("hi", encoding="utf-8")
        with pytest.raises(NotADirectoryError):
            push(bogus, live_xlsm)


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

class TestCLI:
    def test_ls_lists_module_names(
        self, live_xlsm: Path, capsys: pytest.CaptureFixture[str]
    ) -> None:
        from pyopenvba.__main__ import main
        rc = main(["ls", str(live_xlsm)])
        assert rc == 0
        out = capsys.readouterr().out
        with ExcelFile(live_xlsm) as wb:
            for m in wb.vba_project().modules:
                assert m.name in out

    def test_pull_push_cli_round_trip(
        self, live_xlsm: Path, tmp_path: Path
    ) -> None:
        from pyopenvba.__main__ import main
        src_dir = tmp_path / "vba_src"
        work = tmp_path / "work.xlsm"
        shutil.copy(live_xlsm, work)

        assert main(["pull", str(live_xlsm), str(src_dir)]) == 0
        assert main(["push", str(src_dir), str(work)]) == 0

        with ExcelFile(live_xlsm) as a, ExcelFile(work) as b:
            assert a.vba_modules() == b.vba_modules()


# ---------------------------------------------------------------------------
# xlsb / xlam containers
#
# xlsm / xlsb / xlam all share the same OOXML-style ZIP container with
# xl/vbaProject.bin holding the VBA project. The container code path is
# format-agnostic, so re-using the live xlsm bytes under a different
# extension is a valid functional smoke test for the xlsb / xlam paths.
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("ext", [".xlsb", ".xlam"])
class TestOtherZipContainers:
    def test_pull_push_round_trip(
        self, live_xlsm: Path, tmp_path: Path, ext: str
    ) -> None:
        work = tmp_path / f"book{ext}"
        shutil.copy(live_xlsm, work)
        src_dir = tmp_path / "vba_src"
        pull(work, src_dir)
        push(src_dir, work)
        with ExcelFile(live_xlsm) as a, ExcelFile(work) as b:
            assert a.vba_modules() == b.vba_modules()

    def test_set_module_round_trip(
        self, live_xlsm: Path, tmp_path: Path, ext: str
    ) -> None:
        work = tmp_path / f"book{ext}"
        shutil.copy(live_xlsm, work)
        with ExcelFile(work) as wb:
            name = wb.vba_project().modules[0].name
            original = wb.get_module(name)
            wb.set_module(name, original + "\r\n'edited\r\n")
            wb.save()
        with ExcelFile(work) as wb:
            assert "'edited" in wb.get_module(name)
