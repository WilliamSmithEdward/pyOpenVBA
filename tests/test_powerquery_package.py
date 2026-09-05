"""The ZIP inside the blob, and the blob around it.

Excel writes the package with settings of its own -- raw deflate at level
6, a growth-hint extra field, version words 45 and 20 -- so the check
that matters is not "a ZIP that opens" but "the bytes Excel wrote".
Rebuilding each fixture's package from its parts has to give back the
original bytes.
"""

from __future__ import annotations

from pathlib import Path

import pytest

from pyopenvba.exceptions import PowerQueryError
from pyopenvba.powerquery._mashup import DEFAULT_PERMISSIONS, Mashup
from pyopenvba.powerquery._metadata import Metadata
from pyopenvba.powerquery._opc import OpcFile
from pyopenvba.powerquery._package import (
    CONFIG_PART,
    CONTENT_TYPES_PART,
    SECTION_PART,
    Package,
    new_package,
)
from pyopenvba.powerquery.workbook import _split_item  # pyright: ignore[reportPrivateUsage]

FIXTURES = Path(__file__).parent / "fixtures" / "power_query"
WORKBOOKS = sorted(FIXTURES.glob("*.xlsx"))
WITH_QUERIES = [path for path in WORKBOOKS if path.name != "no_queries.xlsx"]


def mashup_of(path: Path) -> Mashup:
    package = OpcFile.parse(path.read_bytes())
    for name in package.names():
        if name.startswith("customXml/item") and name.split("/")[-1].startswith("item") and name.endswith(".xml"):
            raw = package.read(name)
            if b"DataMashup" in raw or b"D\x00a\x00t\x00a\x00M" in raw:
                return Mashup.parse(_split_item(raw)[1])
    raise AssertionError(f"{path.name} carries no mashup")


# --- the inner package --------------------------------------------------------


@pytest.mark.parametrize("path", WITH_QUERIES, ids=lambda p: p.name)
def test_a_package_writes_back_the_bytes_it_was_read_from(path: Path) -> None:
    package = mashup_of(path).package
    assert package.serialize() == package.source


@pytest.mark.parametrize("path", WITH_QUERIES, ids=lambda p: p.name)
def test_a_package_rebuilt_from_its_parts_is_the_one_excel_wrote(path: Path) -> None:
    """Nothing is carried over but the parts and their timestamps, and the
    bytes still come out identical -- which is what says the compression
    settings and the header fields are right."""
    original = mashup_of(path).package
    rebuilt = Package(parts=original.parts, source=None)
    assert rebuilt.serialize() == original.source


@pytest.mark.parametrize("path", WITH_QUERIES, ids=lambda p: p.name)
def test_every_package_holds_the_three_parts_excel_writes(path: Path) -> None:
    package = mashup_of(path).package
    assert [part.name for part in package.parts] == [CONFIG_PART, CONTENT_TYPES_PART, SECTION_PART]
    assert package.read(SECTION_PART).decode("utf-8").startswith("section Section1;")


def test_a_part_can_be_replaced_and_the_rest_stay_as_they_were() -> None:
    package = mashup_of(FIXTURES / "three_queries.xlsx").package
    config = package.read(CONFIG_PART)
    package.write(SECTION_PART, b"section Section1;")
    assert package.read(SECTION_PART) == b"section Section1;"
    assert package.read(CONFIG_PART) == config
    assert package.source is None


def test_writing_a_part_the_same_leaves_the_bytes_alone() -> None:
    package = mashup_of(FIXTURES / "three_queries.xlsx").package
    package.write(SECTION_PART, package.read(SECTION_PART))
    assert package.serialize() == package.source


def test_a_fresh_package_carries_what_excel_carries() -> None:
    package = new_package("section Section1;")
    assert [part.name for part in package.parts] == [CONFIG_PART, CONTENT_TYPES_PART, SECTION_PART]
    assert b"MinVersion" in package.read(CONFIG_PART)
    assert b"application/x-ms-m" in package.read(CONTENT_TYPES_PART)
    assert Package.parse(package.serialize()).read(SECTION_PART) == b"section Section1;"


def test_a_package_that_is_not_one_is_refused() -> None:
    with pytest.raises(PowerQueryError, match="holds no parts"):
        Package.parse(b"not a zip at all")


def test_a_missing_part_is_refused() -> None:
    with pytest.raises(PowerQueryError, match="no part named"):
        new_package("section Section1;").read("Formulas/Nothing.m")


# --- the envelope -------------------------------------------------------------


@pytest.mark.parametrize("path", WITH_QUERIES, ids=lambda p: p.name)
def test_a_blob_writes_back_the_bytes_it_was_read_from(path: Path) -> None:
    mashup = mashup_of(path)
    assert mashup.serialize() == mashup.source


@pytest.mark.parametrize("path", WITH_QUERIES, ids=lambda p: p.name)
def test_every_blob_carries_a_permission_list_and_a_metadata_content(path: Path) -> None:
    """Excel needs both: a blob without them opens with an error rather
    than with its queries."""
    mashup = mashup_of(path)
    assert b"PermissionList" in mashup.permissions
    assert mashup.metadata.content


def test_a_blob_rebuilt_from_its_pieces_matches_what_it_came_from() -> None:
    mashup = mashup_of(FIXTURES / "three_queries.xlsx")
    fresh = Mashup(
        package=Package(parts=mashup.package.parts, source=None),
        metadata=Metadata.parse(mashup.metadata.serialize()),
        permissions=mashup.permissions,
        bindings=mashup.bindings,
        version=mashup.version,
    )
    assert fresh.serialize() == mashup.source


def test_touching_a_blob_drops_the_signature_it_no_longer_matches() -> None:
    """The bindings sign the package as it was; Excel opens and refreshes
    a workbook whose bindings are empty, so they go rather than lie."""
    mashup = mashup_of(FIXTURES / "three_queries.xlsx")
    assert mashup.bindings
    mashup.touch()
    assert mashup.bindings == b""
    assert mashup.source is None


def test_a_blob_that_is_cut_short_is_refused() -> None:
    with pytest.raises(PowerQueryError, match="at least eight bytes"):
        Mashup.parse(b"\x00\x00")
    with pytest.raises(PowerQueryError, match="claims"):
        Mashup.parse(b"\x00\x00\x00\x00" + (999).to_bytes(4, "little") + b"short")


def test_a_default_permission_list_is_the_one_excel_writes() -> None:
    assert b"<CanEvaluateFuturePackages>false</CanEvaluateFuturePackages>" in DEFAULT_PERMISSIONS
    assert b"<FirewallEnabled>true</FirewallEnabled>" in DEFAULT_PERMISSIONS


# --- the workbook container ---------------------------------------------------


@pytest.mark.parametrize("path", WORKBOOKS, ids=lambda p: p.name)
def test_a_workbook_container_writes_back_byte_for_byte(path: Path) -> None:
    raw = path.read_bytes()
    package = OpcFile.parse(raw)
    assert package.serialize() == raw
    package.source = None
    assert package.serialize() == raw


def test_a_part_added_to_a_container_can_be_read_back() -> None:
    package = OpcFile.parse((FIXTURES / "no_queries.xlsx").read_bytes())
    package.write("customXml/item1.xml", b"<hello/>", after="xl/styles.xml")
    assert package.read("customXml/item1.xml") == b"<hello/>"
    assert OpcFile.parse(package.serialize()).read("customXml/item1.xml") == b"<hello/>"
    assert package.names().index("customXml/item1.xml") == package.names().index("xl/styles.xml") + 1


def test_a_container_that_is_not_a_package_is_refused() -> None:
    with pytest.raises(PowerQueryError, match="no ZIP end record"):
        OpcFile.parse(b"nothing here")
