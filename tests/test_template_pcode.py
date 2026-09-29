"""No template the library ships keeps p-code for code its source lacks.

The ZIP-based templates hold no code, so their p-code holds no instruction.
Access's embedded template retains its existing Main() function. Raw-byte
checks also cover unreachable
cache data and unused database pages, which disassembly cannot detect.
Saves formerly kept edited modules' p-code, compiled from their old source,
and the .xlsm template's Module1 had its source blanked that way:
every workbook create_new made carried a comment, TESTING ONLY DO NOT
INCLUDE THIS IN FINAL OUTPUT, that pcodedmp and olevba read back. The
blank files the templates are baked from ship in the wheel as well, so
they are read here too.
"""

from __future__ import annotations

import io
import zipfile
from pathlib import Path

import pytest

from pyopenvba import _templates
from pyopenvba.access import AccessDatabase
from pyopenvba.cfb import CFB
from pyopenvba.vba import parse_vba_project

BLANK_FILES = Path(_templates.__file__).parent / "blank_files"
PROJECT_PART = "vbaProject.bin"
TEMPLATES = sorted(name for name in vars(_templates) if name.startswith("EMPTY_") and name.endswith("_BYTES")
                   and PROJECT_PART.encode() in getattr(_templates, name))
BLANKS = sorted(path.name for path in BLANK_FILES.iterdir() if path.suffix in (".xlsm", ".xlsb", ".xlam", ".docm", ".pptm"))


def _instructions(package_bytes: bytes) -> dict[str, list[str]]:
    """The mnemonics each module's p-code holds, by module name."""
    with zipfile.ZipFile(io.BytesIO(package_bytes)) as package:
        part = next(name for name in package.namelist() if name.endswith(PROJECT_PART))
        project = parse_vba_project(CFB.from_bytes(package.read(part)))
    return {module.name: [instruction.mnemonic for instruction in module.disassemble().iter_instructions()]
            for module in project.modules}


def test_every_project_is_looked_at() -> None:
    assert TEMPLATES == ["EMPTY_DOCM_BYTES", "EMPTY_PPTM_BYTES", "EMPTY_XLAM_BYTES", "EMPTY_XLSB_BYTES",
                         "EMPTY_XLSM_BYTES"]
    assert BLANKS == ["blank_document.docm", "blank_excel_addin.xlam", "blank_presentation.pptm",
                      "blank_workbook.xlsb", "blank_workbook.xlsm"]


@pytest.mark.parametrize("name", TEMPLATES)
def test_no_template_module_holds_p_code(name: str) -> None:
    found = _instructions(getattr(_templates, name))
    assert {module: mnemonics for module, mnemonics in found.items() if mnemonics} == {}


@pytest.mark.parametrize("name", BLANKS)
def test_no_blank_file_module_holds_p_code(name: str) -> None:
    found = _instructions((BLANK_FILES / name).read_bytes())
    assert {module: mnemonics for module, mnemonics in found.items() if mnemonics} == {}


# These removed comments have occurred in shipped templates. Inspect all
# file bytes as well as ZIP members, not just reachable p-code instructions.
REMOVED_COMMENTS = (
    "TESTING ONLY DO NOT INCLUDE THIS IN FINAL OUTPUT",
    "This is a VBA comment line with line number",
)
ALL_BLANKS = sorted(path.name for path in BLANK_FILES.iterdir() if path.is_file())
ALL_TEMPLATES = sorted(name for name in vars(_templates) if name.startswith("EMPTY_") and name.endswith("_BYTES"))
ACCESS_BLANKS = sorted(name for name in ALL_BLANKS if Path(name).suffix in (".accdb", ".mdb"))


def _assert_no_removed_comments(raw: bytes) -> None:
    chunks = [raw]
    if zipfile.is_zipfile(io.BytesIO(raw)):
        with zipfile.ZipFile(io.BytesIO(raw)) as package:
            chunks.extend(package.read(name) for name in package.namelist())
    for chunk in chunks:
        for comment in REMOVED_COMMENTS:
            for encoding in ("ascii", "utf-16le"):
                assert comment.encode(encoding) not in chunk, (comment, encoding)


@pytest.mark.parametrize("name", ALL_BLANKS)
def test_no_shipped_reference_keeps_removed_comment_bytes(name: str) -> None:
    _assert_no_removed_comments((BLANK_FILES / name).read_bytes())


@pytest.mark.parametrize("name", ALL_TEMPLATES)
def test_no_embedded_template_keeps_removed_comment_bytes(name: str) -> None:
    _assert_no_removed_comments(getattr(_templates, name))


def test_access_blank_reference_has_source_only_module() -> None:
    with AccessDatabase((BLANK_FILES / "blank_database.accdb").read_bytes()) as database:
        modules = database.module_streams()
        assert len(modules) == 1
        assert modules[0].name == "Module1"
        assert modules[0].offset == 0
        assert modules[0].text == 'Attribute VB_Name = "Module1"\r\nOption Compare Database\r\n'
        assert not any(name.startswith("__SRP_") for name, _, _ in database.project_streams())


@pytest.mark.parametrize("name", [*ACCESS_BLANKS, "EMPTY_ACCDB_BYTES"])
def test_access_templates_have_no_removed_comments_in_logical_streams(name: str) -> None:
    # Long-value rows can cross page boundaries; scan their assembled bytes too.
    raw = _templates.EMPTY_ACCDB_BYTES if name == "EMPTY_ACCDB_BYTES" else (BLANK_FILES / name).read_bytes()
    with AccessDatabase(raw) as database:
        for _, data, _ in database.project_streams():
            _assert_no_removed_comments(data)


@pytest.mark.parametrize("name", ["engine_skeleton.accdb", "engine_skeleton.mdb"])
def test_access_engine_skeleton_has_no_vba(name: str) -> None:
    with AccessDatabase((BLANK_FILES / name).read_bytes()) as database:
        assert not database.has_vba_project()
        assert not database.module_streams()



def test_cleaning_access_reference_also_removes_freed_page_bytes() -> None:
    # This historical fixture intentionally contains the old cache. It is
    # outside src/ and is excluded from both distribution formats.
    original = AccessDatabase((Path(__file__).parent / "fixtures" / "access_compiled_blank.accdb").read_bytes())
    sources = {module.name: module.source for module in original.modules()}
    for name, source in sources.items():
        original.set_module_source(name, source)
    assert b"This is a VBA comment line with line number" in original.to_bytes()
    compacted = original.compact_and_repair()
    _assert_no_removed_comments(compacted.to_bytes())
    assert {module.name: module.source for module in compacted.modules()} == sources
    assert all(module.offset == 0 for module in compacted.module_streams())
