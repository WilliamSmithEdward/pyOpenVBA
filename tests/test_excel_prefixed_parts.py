"""Workbooks whose SpreadsheetML parts spell their namespace under a prefix, read and saved as Excel does.

tests/fixtures/prefixed_parts/ holds copies of workbooks Excel saved for
other measurements, their parts spelled ``<x:worksheet xmlns:x="...">``,
one with its stylesheet alone so spelled, and Excel's save of each
(scripts/measure_prefixed_parts.py). Excel reads each copy as the workbook
it was made from, and saves every part in its own spelling but a pivot
cache's records.
"""
from __future__ import annotations

import json
import re
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.apps.excel._prefixed_parts import MAIN, unprefixed

FIXTURES = Path(__file__).parent / "fixtures"
FOLDER = FIXTURES / "prefixed_parts"
RECORD: dict[str, Any] = json.loads((FOLDER / "prefixed_parts.json").read_text(encoding="utf-8"))
CASES: list[dict[str, Any]] = RECORD["cases"]


def parts(path: Path) -> dict[str, bytes]:
    with zipfile.ZipFile(path) as package:
        return {name: package.read(name) for name in package.namelist()}


def roots(path: Path) -> dict[str, str]:
    """Each part of SpreadsheetML's under xl/, and its root element as the part spells it, as the record keeps them."""
    out: dict[str, str] = {}
    for part, data in sorted(parts(path).items()):
        if part.startswith("xl/") and part.endswith(".xml"):
            start = re.search(r"<(?![?!])([\w:.-]+)[^>]*>", data.decode("utf-8"))
            if start is not None and MAIN in start.group():
                out[part] = start.group(1)
    return out


def test_the_record_holds_each_file() -> None:
    assert all(case["alike"] for case in CASES)
    assert all(roots(FOLDER / f"{case['name']}.xlsx") == case["prefixed"]
               and roots(FOLDER / f"{case['name']}_resaved.xlsx") == case["resaved"] for case in CASES)


@pytest.mark.parametrize("case", CASES, ids=lambda case: case["name"])
def test_each_copy_spelled_again_is_the_workbook_it_was_made_from(case: dict[str, Any]) -> None:
    """Every part Excel spells its own way comes back to the bytes of the part it was made from; the records of a
    pivot cache, which Excel keeps as the file had them, stay as they are."""
    copy, source = parts(FOLDER / f"{case['name']}.xlsx"), parts(FIXTURES / case["source"])
    assert copy.keys() == source.keys()
    for name, data in copy.items():
        if not (name.startswith("xl/") and name.endswith(".xml")):
            assert data == source[name], name
            continue
        text = data.decode("utf-8")
        if case["prefixed"].get(name, "").startswith("x:") and not case["resaved"][name].startswith("x:"):
            assert unprefixed(text).encode("utf-8") == source[name], name
        else:
            assert unprefixed(text) == text, name


@pytest.mark.parametrize("case", CASES, ids=lambda case: case["name"])
def test_a_copy_reads_as_excel_reads_it(case: dict[str, Any]) -> None:
    """Every cell's formula, number format, bold and fill, each sheet's rules and tables, as Excel read them."""
    app = ExcelApplication()
    app.add_module(RECORD["read"], name="Reader")
    assert app.run("Read", str(FOLDER / f"{case['name']}.xlsx"), "") == case["answer"]


@pytest.mark.parametrize("case", CASES, ids=lambda case: case["name"])
def test_a_copy_saves_as_the_workbook_it_was_made_from(case: dict[str, Any], tmp_path: Path) -> None:
    """Saved with nothing changed, a copy holds what the workbook it was made from holds saved the same way, each
    part spelled as Excel's save spells it: the namespace the default, but in a pivot cache's records."""
    copy = ExcelApplication.open(FOLDER / f"{case['name']}.xlsx", with_vba=False).save(tmp_path / "copy.xlsx")
    source = ExcelApplication.open(FIXTURES / case["source"], with_vba=False).save(tmp_path / "source.xlsx")
    assert roots(copy) == case["resaved"]
    held, expected = parts(copy), parts(source)
    for name in case["resaved"]:
        if case["resaved"][name].startswith("x:"):
            expected[name] = parts(FOLDER / f"{case['name']}.xlsx")[name]
    assert held == expected


@pytest.mark.parametrize("text", [
    f'<worksheet xmlns="{MAIN}"><sheetData/></worksheet>',
    f'<x:pivotCacheRecords xmlns:x="{MAIN}" count="0"/>',
    f'<x:worksheet xmlns:x="{MAIN}" xmlns="urn:other"><x:sheetData/><a/></x:worksheet>',
    f'<x:worksheet xmlns:x="{MAIN}"><x:sheetData x:extra="1"/></x:worksheet>',
    f'<x:worksheet xmlns:x="{MAIN}"><x:sheetData/><other/></x:worksheet>',
    f'<x:worksheet xmlns:x="{MAIN}"><x:sheetData><x:row xmlns:x="urn:other"/></x:sheetData></x:worksheet>',
    f'<x:worksheet xmlns:x="{MAIN}" xmlns:mc="urn:mc" mc:Ignorable="x"><x:sheetData/></x:worksheet>',
], ids=["unprefixed", "records", "default-namespace", "prefixed-attribute", "no-namespace", "redeclared",
        "ignorable"])
def test_a_part_that_would_read_otherwise_is_left_as_it_is(text: str) -> None:
    assert unprefixed(text) == text


def test_markup_that_names_no_element_is_left_as_it_is() -> None:
    """A comment, CDATA and the XML declaration keep their text, and the declaration keeps its place and quote."""
    text = (f"<?xml version=\"1.0\"?>\n<x:sst\n  xmlns:x='{MAIN}' count=\"1\"><!-- <x:si> --><x:si><x:t>"
            "<![CDATA[<x:t>]]></x:t></x:si></x:sst>")
    assert unprefixed(text) == (f"<?xml version=\"1.0\"?>\n<sst\n  xmlns=\"{MAIN}\" count=\"1\"><!-- <x:si> --><si>"
                                "<t><![CDATA[<x:t>]]></t></si></sst>")
