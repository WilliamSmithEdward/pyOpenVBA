"""Shared reference edits across file containers, including reference-only saves."""
from pathlib import Path

import pytest

from pyopenvba import ExcelFile, WordFile, PowerPointFile
from pyopenvba.access import AccessDatabase
from pyopenvba._references import ReferenceManager, reference_spans
from pyopenvba.exceptions import VBAProjectError


@pytest.mark.parametrize("kind,suffix,library", [
    (ExcelFile, ".xlsm", "Word"), (WordFile, ".docm", "Excel"),
    (PowerPointFile, ".pptm", "Excel"), (AccessDatabase, ".accdb", "Excel"),
])
def test_shared_reference_surface(kind: type[ExcelFile] | type[WordFile] | type[PowerPointFile] | type[AccessDatabase],
                                  suffix: str, library: str, tmp_path: Path) -> None:
    path = tmp_path / ("references" + suffix)
    file = kind.create_new(path)
    before = [(r.name, r.libid) for r in file.references()]
    added = file.add_reference(library)
    assert added.kind == "registered"
    assert file.add_reference(library.lower()).guid == added.guid
    assert len(file.references()) == len(before) + 1
    file.save(path)
    file = kind(path)
    assert file.references()[-1].guid == added.guid
    assert file.remove_reference(added.guid.strip("{}"))
    assert not file.remove_reference(library)
    file.save(path)
    assert [(r.name, r.libid) for r in kind(path).references()] == before
    assert kind.add_reference is ReferenceManager.add_reference
    assert kind.remove_reference is ReferenceManager.remove_reference


def test_reference_only_write_invalidates_cache_and_preserves_modules(tmp_path: Path) -> None:
    path = tmp_path / "references.xlsm"
    file = ExcelFile.create_new(path)
    before = file.vba_modules()
    raw = file.vba_project().dir_raw
    at = file.vba_project().dir_modules_offset
    file.add_reference("Word")
    file.save(path)
    with ExcelFile(path) as reopened:
        assert reopened.vba_modules() == before
        project = reopened.vba_project()
        assert project.dir_raw[project.dir_modules_offset:] == raw[at:]
        # The existing cache invalidator clears compiled state.
        from pyopenvba.cfb import CFB
        cfb = CFB.from_bytes(reopened.vba_project_bytes())
        assert not any(cfb.get_stream_in_storage("VBA", "_VBA_PROJECT")[5:])


def test_alternating_module_and_reference_saves(tmp_path: Path) -> None:
    path = tmp_path / "mixed.xlsm"
    file = ExcelFile.create_new(path)
    file.add_reference("Word")
    file.vba_project().add_module("Extra", "Public Function Answer() As Long\nAnswer = 42\nEnd Function")
    file.save(path)
    file.remove_reference("Word")
    file.save(path)
    with ExcelFile(path) as reopened:
        assert "Extra" in reopened.module_names()
        assert "42" in reopened.get_module("Extra")
        assert all(r.name != "Word" for r in reopened.references())


def test_word_removal_preserves_normal_and_other_bytes(tmp_path: Path) -> None:
    source = Path(__file__).parent / "fixtures/WordExcelInteropFixture.docm"
    file = WordFile(source)
    raw = file.vba_project().dir_raw
    spans = reference_spans(raw)
    normal = next(s for s in spans if s.reference.kind == "project")
    target = next(s for s in spans if s.reference.name == "Excel")
    assert file.remove_reference("Excel")
    expected = raw[:target.start] + raw[target.end:]
    assert file.vba_project().dir_raw == expected
    path = tmp_path / "removed.docm"
    file.save(path)
    with WordFile(path) as reopened:
        assert reopened.vba_project().dir_raw == expected
        assert any(r.name == normal.reference.name for r in reopened.references())


def test_msforms_removal_protects_designer() -> None:
    source = Path(__file__).parent / "live_excel_testing/nested_form.xlsm"
    with ExcelFile(source) as file:
        original = file.vba_project().dir_raw
        with pytest.raises(VBAProjectError, match="UserForms"):
            file.remove_reference("MSForms")
        assert file.vba_project().dir_raw == original


def test_control_reference_span_includes_embedded_names() -> None:
    source = Path(__file__).parent / "live_excel_testing/nested_form.xlsm"
    with ExcelFile(source) as file:
        raw = file.vba_project().dir_raw
        span = next(s for s in reference_spans(raw) if s.reference.name == "MSForms")
        assert span.reference.kind == "control"
        remaining = raw[:span.start] + raw[span.end:]
        assert all(s.reference.name != "MSForms" for s in reference_spans(remaining))
        assert len(reference_spans(remaining)) == len(reference_spans(raw)) - 1


def test_invalid_and_implicit_edits_are_atomic(tmp_path: Path) -> None:
    file = ExcelFile.create_new(tmp_path / "invalid.xlsm")
    raw = file.vba_project().dir_raw
    for name in ("Excel", "VBA", "Unknown"):
        with pytest.raises(VBAProjectError):
            file.add_reference(name)
    with pytest.raises(VBAProjectError):
        file.add_reference("Custom", "not-a-guid")
    with pytest.raises(VBAProjectError):
        file.remove_reference("")
    assert file.vba_project().dir_raw == raw
    assert not file.vba_project().dir_references_dirty


def test_remove_complete_control_group_without_a_designer(tmp_path: Path) -> None:
    fixture = Path(__file__).parent / "live_excel_testing/nested_form.xlsm"
    with ExcelFile(fixture) as source:
        raw = source.vba_project().dir_raw
        span = next(s for s in reference_spans(raw) if s.reference.name == "MSForms")
        block = raw[span.start:span.end]
    path = tmp_path / "control_reference.xlsm"
    file = ExcelFile.create_new(path)
    project = file.vba_project()
    original = project.dir_raw
    at = project.dir_modules_offset
    project.dir_raw = original[:at] + block + original[at:]
    project.dir_modules_offset += len(block)
    assert file.remove_reference("MSForms")
    assert project.dir_raw == original
    file.save()
    with ExcelFile(path) as reopened:
        assert reopened.vba_project().dir_raw == original


def test_reference_only_edits_honor_project_protection(tmp_path: Path) -> None:
    fixture = Path(__file__).parent / "live_excel_testing/workbook_with_password_protected_vba_modules.xlsm"
    with ExcelFile(fixture) as file:
        file.add_reference("Word")
        with pytest.raises(VBAProjectError, match="password-protected"):
            file.save(tmp_path / "protected.xlsm")
        assert file.vba_project().dir_references_dirty
