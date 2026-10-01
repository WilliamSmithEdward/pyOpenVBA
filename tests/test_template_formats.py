"""Templates and add-ins that share a supported container (issue #39).

.xltm, .ppsm and .ppam are the .xlsm and .pptm packages; .xlt, .xla and .dot
are the .xls and .doc compound files. They open and read as those do, and
refuse to save, since no writer for them has been measured.
"""

from __future__ import annotations

import shutil
from pathlib import Path

import pytest

from pyopenvba import ExcelFile, PowerPointFile, WordFile
from pyopenvba._host import VBAHostFile
from pyopenvba.exceptions import UnsupportedFormatError

FIXTURES = Path(__file__).parent / "fixtures"

CASES: list[tuple[type[VBAHostFile], str, str]] = [
    (ExcelFile, "first_macro/excel_module.xlsm", ".xltm"),
    (ExcelFile, "binary_project/workbook.xls", ".xlt"),
    (ExcelFile, "binary_project/workbook.xls", ".xla"),
    (WordFile, "binary_project/document.doc", ".dot"),
    (PowerPointFile, "first_macro/powerpoint_module.pptm", ".ppsm"),
    (PowerPointFile, "first_macro/powerpoint_module.pptm", ".ppam"),
]
IDS = [suffix for _, _, suffix in CASES]


def as_suffix(tmp_path: Path, source: str, suffix: str) -> Path:
    path = tmp_path / f"file{suffix}"
    shutil.copyfile(FIXTURES / source, path)
    return path


@pytest.mark.parametrize(("host", "source", "suffix"), CASES, ids=IDS)
def test_reads_the_modules_of_the_container_it_shares(
    tmp_path: Path, host: type[VBAHostFile], source: str, suffix: str
) -> None:
    with host(FIXTURES / source) as original:
        expected = original.vba_modules()
    assert expected
    with host(as_suffix(tmp_path, source, suffix)) as copy:
        assert copy.vba_modules() == expected


@pytest.mark.parametrize(("host", "source", "suffix"), CASES, ids=IDS)
def test_refuses_to_save(tmp_path: Path, host: type[VBAHostFile], source: str, suffix: str) -> None:
    path = as_suffix(tmp_path, source, suffix)
    before = path.read_bytes()
    with host(path) as copy, pytest.raises(UnsupportedFormatError, match="read only"):
        copy.save()
    assert path.read_bytes() == before
