"""Edited module caches must not retain removed text (#35), across host containers."""
from __future__ import annotations

import os
from pathlib import Path
from typing import Any

import pytest

from pyopenvba import ExcelFile, PowerPointFile, WordFile
from pyopenvba.vba import VBAModuleKind

ROOT = Path(__file__).parent
CASES: list[tuple[Any, str, str, str]] = [
    (ExcelFile, "live_excel_testing/test_macro_workbook.xlsm", "ExcelSession", "EXCEL"),
    (ExcelFile, "live_excel_testing/test_macro_workbook.xlsb", "ExcelSession", "EXCEL"),
    (WordFile, "live_word_testing/simple_macros.docm", "WordSession", "WORD"),
    (PowerPointFile, "live_powerpoint_testing/simple_macros.pptm", "PowerPointSession", "POWERPOINT"),
    (ExcelFile, "fixtures/binary_project/workbook.xls", "ExcelSession", "EXCEL"),
    (WordFile, "fixtures/binary_project/document.doc", "WordSession", "WORD"),
]


@pytest.mark.parametrize(("opener", "fixture", "session", "host"), CASES)
def test_edited_cache_removed_and_repeated_saves(opener: Any, fixture: str, session: str, host: str, tmp_path: Path) -> None:
    path = ROOT / fixture
    out = tmp_path / path.name
    with opener(path) as file:
        project = file.vba_project()
        module = next(m for m in project.modules if m.kind == VBAModuleKind.standard)
        assert module.prefix_bytes, "fixture must contain a compiled module"
        before = {m.name: (m.source, m.prefix_bytes, m.text_offset) for m in project.modules}
        name = module.name
        source = f'Attribute VB_Name = "{name}"\r\nPublic Function CleanProbe() As Long\r\nCleanProbe = 731\r\nEnd Function\r\n'
        file.set_module(name, source)
        file.save(out)
        assert module.text_offset == 0 and module.prefix_bytes == b""
        assert not list(module.disassemble().iter_instructions())
        file.save(out)  # Metadata must also be correct on the same object's next save.
    with opener(out) as file:
        for m in file.vba_project().modules:
            if m.name == name:
                assert m.source == source
                assert m.text_offset == 0 and m.prefix_bytes == b""
                assert not list(m.disassemble().iter_instructions())
            else:
                assert (m.source, m.prefix_bytes, m.text_offset) == before[m.name]
        file.set_module(name, source.replace("731", "732"))
        file.save(out)
    with opener(out) as file:
        assert "732" in file.get_module(name)
        assert file.vba_project().get_module(name).text_offset == 0


@pytest.mark.parametrize(("opener", "fixture", "session", "host"), CASES)
def test_live_office_runs_edited_source_without_old_cache(opener: Any, fixture: str, session: str, host: str, tmp_path: Path) -> None:
    if os.environ.get(f"RUN_LIVE_{host}") != "1":
        pytest.skip(f"set RUN_LIVE_{host}=1 for the live Office gate")
    harness = pytest.importorskip("pyvbaharness")
    path = tmp_path / ("native" + Path(fixture).suffix)
    out = tmp_path / ("edited" + path.suffix)
    config = harness.HarnessConfig(lock_wait_s=120.0)
    old_source = ("' Internal note: remove before shipping\r\n"
                  "Public Function CleanProbe() As String\r\n"
                  'CleanProbe = "example-secret-7731"\r\nEnd Function\r\n')
    with getattr(harness, session)(config) as office:
        office.new_document()
        result = office.run_vba(old_source, "CleanProbe", module_name="PcodeProbe", timeout=60.0)
        assert result.ok, f"{result.outcome}: {result.message}"
        assert result.value == "example-secret-7731"
        if path.suffix in (".xls", ".doc"):
            save = (f'ThisWorkbook.SaveAs "{path}", 56' if path.suffix == ".xls"
                    else f'ActiveDocument.SaveAs2 "{path}", 0')
            saved = office.run_vba(f"Public Sub SaveNative()\n{save}\nEnd Sub", "SaveNative", timeout=60.0)
            assert saved.ok, f"{saved.outcome}: {saved.message}"
        else:
            office.save_as(path)
    with opener(path) as file:
        module = file.vba_project().get_module("PcodeProbe")
        opcodes = {i.mnemonic for i in module.disassemble().iter_instructions()}
        assert {"QuoteRem", "LitStr"} <= opcodes
        file.set_module("PcodeProbe", 'Public Function CleanProbe() As Long\r\nCleanProbe = 731\r\nEnd Function\r\n')
        file.save(out)
    with opener(out) as file:
        assert not list(file.vba_project().get_module("PcodeProbe").disassemble().iter_instructions())
        project_bytes = file.vba_project_bytes()
        for removed in ("example-secret-7731", "Internal note: remove before shipping"):
            assert removed.encode("ascii") not in project_bytes
            assert removed.encode("utf-16le") not in project_bytes
    with getattr(harness, session)(config) as office:
        office.open_document(out, read_only=False)
        result = office.run_macro("PcodeProbe.CleanProbe", timeout=60.0)
        assert result.ok, f"{result.outcome}: {result.message} {result.dialogs}"
        assert result.value == 731
