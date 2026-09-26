"""The order Office writes a part's relationships in, replayed against live Excel.

tests/fixtures/relationships/orders.json is what
scripts/measure_relationship_order.py saw: the order Excel wrote the
relationships of a sheet holding n hyperlinks in, for every n from 1 to
160 and for 17 larger n up to 65,530, the most a sheet holds.
pyopenvba._relationships puts the same ids in the same order from their
hash, and leaves a part whose order it cannot vouch for as it is.
"""

from __future__ import annotations

import json
import re
import zipfile
from pathlib import Path

import pytest

from pyopenvba._relationships import bucket_order, id_hash, in_office_order

FIXTURES = Path(__file__).parent / "fixtures"
ORDERS: dict[str, list[int]] = json.loads((FIXTURES / "relationships" / "orders.json").read_text("utf-8"))["orders"]
_HEAD = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
         '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">')


def _part(numbers: list[int]) -> str:
    return _HEAD + "".join(f'<Relationship Id="rId{number}" Type="t" Target="x{number}"/>'
                           for number in numbers) + "</Relationships>"


@pytest.mark.parametrize("count", list(ORDERS))
def test_ids_come_in_excels_order(count: str) -> None:
    assert bucket_order(list(range(1, int(count) + 1))) == ORDERS[count]


@pytest.mark.parametrize("count", ["2", "7", "33", "128", "160", "300", "65530"])
def test_a_part_is_put_in_excels_order(count: str) -> None:
    assert in_office_order(_part(list(range(1, int(count) + 1)))) == _part(ORDERS[count])


def test_the_hash_is_base_101_through_a_step_of_rand() -> None:
    # "rId1" read in base 101 is 118,209,136; one step of C's rand() generator keeps bits 16 to 31 of the next.
    assert sum(ord(one) * 101 ** place for place, one in enumerate(reversed("rId1"))) == 118209136
    assert id_hash("rId1") == ((118209136 * 1103515245 + 12345) % 2**32) >> 16 == 13676


def test_ids_with_gaps_are_ordered_by_their_hash() -> None:
    # A part the model wrote after taking one away: Office's table places the rest by their own hashes.
    numbers = [1, 2, 4, 5, 9]
    ordered = bucket_order(numbers)
    assert ordered is not None and sorted(ordered) == numbers
    assert ordered == sorted(numbers, key=lambda number: (id_hash(f"rId{number}") % 8, number))


@pytest.mark.parametrize("numbers", [list(range(1, 65538)), [1, 1]], ids=["65,537 ids", "twice"])
def test_a_part_whose_order_is_not_known_is_left_as_it_is(numbers: list[int]) -> None:
    assert in_office_order(_part(numbers)) == _part(numbers)


@pytest.mark.parametrize("first", ["R1", "rId01"])
def test_a_part_with_ids_of_another_form_is_left_as_it_is(first: str) -> None:
    text = _HEAD + f'<Relationship Id="{first}" Type="t" Target="a"/><Relationship Id="rId2" Type="t" Target="b"/>' \
        "</Relationships>"
    assert in_office_order(text) == text


@pytest.mark.parametrize("workbook", ["notes.xlsx", "row_formats/row_formats.xlsx", "formula_corpus/inputs.xlsx"])
def test_the_parts_excel_wrote_are_in_its_order(workbook: str) -> None:
    with zipfile.ZipFile(FIXTURES / workbook) as package:
        parts = [package.read(name).decode("utf-8") for name in package.namelist() if name.endswith(".rels")]
    assert parts
    for text in parts:
        assert re.findall(r"<Relationship\b", text)
        assert in_office_order(text) == text
