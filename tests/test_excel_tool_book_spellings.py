"""A book spelled as a writer may spell it: each vocabulary under another prefix, and what may be left out for its default
(scripts/measure_tool_book_spellings.py). Excel read every variant; the model reads each as Excel did."""

from __future__ import annotations

import importlib.util
import json
import sys
from pathlib import Path
from types import ModuleType
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.apps.excel._prefixed_parts import addressed, prefixes_spelled, rels_unprefixed
from test_excel_tool_books import READS

HERE = Path(__file__).parent / "fixtures" / "tool_books"
RECORD: dict[str, Any] = json.loads((HERE / "spellings.json").read_text(encoding="utf-8"))
ROWS: dict[str, Any] = RECORD["variants"]
#: Excel gives these 1004 at open and the model opens them. Which detail refuses each was not split: a sheet with a table, its
#: cells or rows unnumbered, and a stylesheet with no cellStyleXfs.
EXCEL_REFUSES = {"no_cell_addresses_by_a_table", "no_row_numbers_by_a_table", "no_cell_style_xfs"}


def _script() -> ModuleType:
    folder = Path(__file__).parents[1] / "scripts"
    sys.path.insert(0, str(folder))
    spec = importlib.util.spec_from_file_location("measure_tool_book_spellings", folder / "measure_tool_book_spellings.py")
    assert spec is not None and spec.loader is not None
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


SCRIPT = _script()


@pytest.mark.parametrize("variant", [name for name, row in ROWS.items() if row["excel"]])
def test_reads_what_excel_reads(variant: str, tmp_path: Path) -> None:
    """Every read in one run, where nothing stops a run: a list of differences to read rather than a test for each."""
    file = tmp_path / f"{variant}.xlsx"
    SCRIPT.made(HERE / "xlsxwriter.xlsx", variant, file)
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(SCRIPT.source(READS, "typed"), name="M")
    answers = str(app.run("Go", file.as_posix())).split("\t")
    assert [(read, theirs, ours) for read, theirs, ours in zip(READS, ROWS[variant]["excel"], answers, strict=True) if theirs != ours] == []


def test_every_variant_was_measured_and_the_refusals_are_the_named_ones() -> None:
    assert set(ROWS) == set(SCRIPT.VARIANTS)
    assert {name for name, row in ROWS.items() if not row["excel"]} == EXCEL_REFUSES


def test_cells_and_rows_with_no_address_go_by_position() -> None:
    given = b'<sheetData><row><c><v>1</v></c><c t="s"><v>2</v></c></row><row r="5"><c r="C5"><v>3</v></c><c/></row><row><c/></row></sheetData>'
    assert addressed(given) == (b'<sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1" t="s"><v>2</v></c></row>'
                                b'<row r="5"><c r="C5"><v>3</v></c><c r="D5"/></row><row r="6"><c r="A6"/></row></sheetData>')
    whole = b'<sheetData><row r="1"><c r="A1"/></row></sheetData>'
    assert addressed(whole) is whole


def test_a_prefix_of_its_own_is_respelled_for_the_readers() -> None:
    uri = b"http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    assert prefixes_spelled(b'<worksheet xmlns:rel="' + uri + b'"><drawing rel:id="rId1"/></worksheet>') == (
        b'<worksheet xmlns:r="' + uri + b'"><drawing r:id="rId1"/></worksheet>')
    usual = b'<worksheet xmlns:r="' + uri + b'"><drawing r:id="rId1"/></worksheet>'
    assert prefixes_spelled(usual) is usual


def test_a_relationships_part_under_a_prefix_is_read_as_any_other() -> None:
    assert rels_unprefixed(b'<p:Relationships xmlns:p="u"><p:Relationship Id="a"/></p:Relationships>') == (
        b'<Relationships xmlns="u"><Relationship Id="a"/></Relationships>')
