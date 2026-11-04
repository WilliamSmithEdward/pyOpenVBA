"""Every workbook kept as a fixture has well-formed document properties, or opening it would be error 1004 where Excel saved it.

A mask once wrote the author as a bare <author> into 41 of them (scripts/measure_document_properties.py).
"""

from __future__ import annotations

import zipfile
from pathlib import Path
from xml.parsers import expat

FIXTURES = Path(__file__).parent / "fixtures"
BOOKS = {".xlsx", ".xlsm", ".xltx", ".xltm", ".xlam"}


def test_no_workbook_fixture_has_properties_that_are_not_xml() -> None:
    malformed: list[str] = []
    for path in FIXTURES.rglob("*"):
        if path.suffix.lower() not in BOOKS or not zipfile.is_zipfile(path):
            continue
        with zipfile.ZipFile(path) as package:
            for name in package.namelist():
                if name.startswith("docProps/"):
                    try:
                        expat.ParserCreate().Parse(package.read(name), True)
                    except expat.ExpatError as failure:
                        malformed.append(f"{path.relative_to(FIXTURES).as_posix()} {name}: {failure}")
    assert malformed == []
