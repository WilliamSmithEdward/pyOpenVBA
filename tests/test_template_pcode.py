"""No template the library ships keeps p-code for code its source lacks.

A template's modules hold no code, so their p-code holds no instruction.
A save keeps an edited module's p-code, compiled from the source it had
before, and the .xlsm template's Module1 had its source blanked that way:
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
