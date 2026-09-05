"""The metadata a Power Query package carries beside its M.

The expected values here are not hand-written.  Each one came out of
Microsoft's own packaging assemblies, which ship with Excel, driven
through a small PowerShell oracle: ``SerializedMetadataEntry`` for the
value encodings, ``SerializedPackageItemLocation.ItemPathFromParts`` for
the paths, ``QueriesMetadataSerializer.SerializeQueryGroups`` for the
groups and ``PackageMetadataSerializer.Serialize`` for the section as a
whole.  A change here that these no longer match is a change Excel would
not have written.
"""

from __future__ import annotations

import base64
import datetime as dt
import uuid

import pytest

from pyopenvba.exceptions import PowerQueryError
from pyopenvba.powerquery import _metadata as meta
from pyopenvba.powerquery._metadata import Entry, Item, Metadata, QueryGroup

GROUP_A = uuid.UUID("aaaaaaaa-1111-2222-3333-444444444444")
GROUP_B = uuid.UUID("bbbbbbbb-1111-2222-3333-444444444444")

#: (groups, what Microsoft's serializer produced for them)
GROUP_CASES: list[tuple[list[QueryGroup], str]] = [
    ([], "AAAAAA=="),
    (
        [QueryGroup(GROUP_A, "Staging")],
        "AQAAAAAAAACqqqqqEREiIjMzREREREREB1N0YWdpbmcAAAAAAAA=",
    ),
    (
        [QueryGroup(GROUP_A, "Staging", "notes here")],
        "AQAAAAAAAACqqqqqEREiIjMzREREREREB1N0YWdpbmcKbm90ZXMgaGVyZQAAAAAA",
    ),
    (
        [QueryGroup(GROUP_A, "Top"), QueryGroup(GROUP_B, "Child", "", GROUP_A, 1)],
        "AgAAAAAAAACqqqqqEREiIjMzREREREREA1RvcAAAAAAAAAAAAAC7u7u7EREiIjMzREREREREBUNoaWxkAAGqqqqqEREiIjMzREREREREAQAAAA==",
    ),
    (
        [QueryGroup(GROUP_A, "Café 日本")],
        "AQAAAAAAAACqqqqqEREiIjMzREREREREDENhZsOpIOaXpeacrAAAAAAAAA==",
    ),
    (
        [QueryGroup(GROUP_A, "N" * 130, "", None, 258)],
        "AQAAAAAAAACqqqqqEREiIjMzREREREREggFOTk5OTk5OTk5OTk5OTk5OTk5OTk5OTk5OTk5OTk5OTk5OTk5OTk5OTk5OTk5OTk5O"
        "Tk5OTk5OTk5OTk5OTk5OTk5OTk5OTk5OTk5OTk5OTk5OTk5OTk5OTk5OTk5OTk5OTk5OTk5OTk5OTk5OTk5OTk5OTk5OTk5OTk5O"
        "Tk5OTk5OAAACAQAA"
    ),
    (
        [QueryGroup(GROUP_A, "X", "", None, -5)],
        "AQAAAAAAAACqqqqqEREiIjMzREREREREAVgAAPv///8=",
    ),
]

#: (part, the path Microsoft's helper produced for it)
PATH_CASES = [
    ("Plain", "Plain"),
    ("With Space", "With%20Space"),
    ("With%Pct", "With%25Pct"),
    ("a/b", "a%2Fb"),
    ("Café", "Caf%C3%A9"),
    ("日本", "%E6%97%A5%E6%9C%AC"),
    ("q'x", "q'x"),
    ("we<ird>", "we%3Cird%3E"),
    ("t\tab", "t%09ab"),
    ("+plus", "%2Bplus"),
]

#: The whole section Microsoft's serializer wrote for the items below.
METADATA_SECTION = base64.b64decode(
    "AAAAACcDAADvu788P3htbCB2ZXJzaW9uPSIxLjAiIGVuY29kaW5nPSJ1dGYtOCI/PjxMb2NhbFBhY2thZ2VNZXRhZGF0YUZpbGUgeG1sbnM6"
    "eHNkPSJodHRwOi8vd3d3LnczLm9yZy8yMDAxL1hNTFNjaGVtYSIgeG1sbnM6eHNpPSJodHRwOi8vd3d3LnczLm9yZy8yMDAxL1hNTFNjaGVt"
    "YS1pbnN0YW5jZSI+PEl0ZW1zPjxJdGVtPjxJdGVtTG9jYXRpb24+PEl0ZW1UeXBlPkFsbEZvcm11bGFzPC9JdGVtVHlwZT48SXRlbVBhdGgg"
    "Lz48L0l0ZW1Mb2NhdGlvbj48U3RhYmxlRW50cmllcyAvPjwvSXRlbT48SXRlbT48SXRlbUxvY2F0aW9uPjxJdGVtVHlwZT5Gb3JtdWxhPC9J"
    "dGVtVHlwZT48SXRlbVBhdGg+U2VjdGlvbjEvV2l0aCUyMFNwYWNlPC9JdGVtUGF0aD48L0l0ZW1Mb2NhdGlvbj48U3RhYmxlRW50cmllcz48"
    "RW50cnkgVHlwZT0iSXNQcml2YXRlIiBWYWx1ZT0ibDAiIC8+PEVudHJ5IFR5cGU9IkZpbGxPYmplY3RUeXBlIiBWYWx1ZT0ic0Nvbm5lY3Rp"
    "b25Pbmx5IiAvPjxFbnRyeSBUeXBlPSJRdWVyeUlEIiBWYWx1ZT0iczQ3YmMyZGFhLWQ4MDEtNDI0Yi1iZDBlLTQ0ZDNmNjdmZTMzNCIgLz48"
    "RW50cnkgVHlwZT0iT2RkIiBWYWx1ZT0ic2EmbHQ7YiZndDsmYW1wO2MmcXVvdDtkJ2UiIC8+PC9TdGFibGVFbnRyaWVzPjwvSXRlbT48SXRl"
    "bT48SXRlbUxvY2F0aW9uPjxJdGVtVHlwZT5Gb3JtdWxhPC9JdGVtVHlwZT48SXRlbVBhdGg+U2VjdGlvbjEvV2l0aCUyMFNwYWNlL1NvdXJj"
    "ZTwvSXRlbVBhdGg+PC9JdGVtTG9jYXRpb24+PFN0YWJsZUVudHJpZXMgLz48L0l0ZW0+PC9JdGVtcz48L0xvY2FsUGFja2FnZU1ldGFkYXRh"
    "RmlsZT4WAAAAUEsFBgAAAAAAAAAAAAAAAAAAAAAAAA=="
)


def section_items() -> list[Item]:
    return [
        Item(meta.ITEM_ALL_FORMULAS, ""),
        Item(
            meta.ITEM_FORMULA,
            meta.item_path(meta.SECTION, "With Space"),
            [
                Entry.of_int(meta.IS_PRIVATE, 0),
                Entry.of_text(meta.FILL_OBJECT_TYPE, meta.FILL_CONNECTION_ONLY),
                Entry.of_text(meta.QUERY_ID, "47bc2daa-d801-424b-bd0e-44d3f67fe334"),
                Entry.of_text("Odd", 'a<b>&c"d\'e'),
            ],
        ),
        Item(meta.ITEM_FORMULA, meta.item_path(meta.SECTION, "With Space", "Source")),
    ]


# --- values -------------------------------------------------------------------


@pytest.mark.parametrize(
    ("entry", "raw"),
    [
        (Entry.of_int(meta.IS_PRIVATE, 0), "l0"),
        (Entry.of_int(meta.FILL_COUNT, -7), "l-7"),
        (Entry.of_flag(meta.FILL_ENABLED, True), "l1"),
        (Entry.of_flag(meta.FILL_ENABLED, False), "l0"),
        (Entry.of_double("D", 0.0), "f0"),
        (Entry.of_double("D", 1 / 3), "f0.333333333333333"),
        (Entry.of_double("D", 1e20), "f1E+20"),
        (Entry.of_text(meta.FILL_OBJECT_TYPE, "ConnectionOnly"), "sConnectionOnly"),
        (Entry.of_guid(meta.QUERY_ID, GROUP_A), "caaaaaaaa-1111-2222-3333-444444444444"),
        (
            Entry.of_time(
                meta.FILL_LAST_UPDATED,
                dt.datetime(2026, 9, 5, 13, 45, 12, 250000, tzinfo=dt.timezone.utc),
            ),
            "d2026-09-05T13:45:12.2500000Z",
        ),
    ],
)
def test_a_value_carries_its_type_the_way_the_format_carries_it(entry: Entry, raw: str) -> None:
    assert entry.raw == raw


def test_a_value_reads_back_as_what_it_stands_for() -> None:
    assert Entry("K", "l42").value == 42
    assert Entry("K", "l1").flag is True
    assert Entry("K", "l0").flag is False
    assert Entry("K", "f1.5").value == 1.5
    assert Entry("K", "sTable").value == "Table"
    assert Entry("K", f"c{GROUP_A}").value == GROUP_A
    assert Entry("K", "d2026-09-05T13:45:12.2500000Z").value == dt.datetime(
        2026, 9, 5, 13, 45, 12, 250000, tzinfo=dt.timezone.utc
    )


def test_a_value_of_a_type_this_format_does_not_define_stays_as_it_was() -> None:
    """Reading must not lose a value it does not recognise: the entry
    would be written back wrong."""
    assert Entry("K", "?strange").value == "?strange"


@pytest.mark.parametrize(("value", "text"), [(0.0, "0"), (-0.0, "0"), (1.5, "1.5"), (1e-7, "1E-07")])
def test_a_double_is_spelled_the_way_dot_net_spells_it(value: float, text: str) -> None:
    assert meta.format_double(value) == text


# --- paths --------------------------------------------------------------------


@pytest.mark.parametrize(("part", "escaped"), PATH_CASES)
def test_an_item_path_escapes_what_the_format_escapes(part: str, escaped: str) -> None:
    assert meta.escape_path_part(part) == escaped
    assert meta.unescape_path_part(escaped) == part


def test_a_path_is_built_and_read_part_by_part() -> None:
    path = meta.item_path("Section1", "With Space", "Changed Type")
    assert path == "Section1/With%20Space/Changed%20Type"
    assert meta.path_parts(path) == ("Section1", "With Space", "Changed Type")


def test_a_broken_escape_is_refused() -> None:
    with pytest.raises(PowerQueryError, match="percent escape"):
        meta.unescape_path_part("a%zz")


# --- query groups -------------------------------------------------------------


@pytest.mark.parametrize(("groups", "value"), GROUP_CASES)
def test_groups_serialize_as_microsofts_own_serializer_serializes_them(
    groups: list[QueryGroup], value: str
) -> None:
    assert meta.pack_groups(groups) == value


@pytest.mark.parametrize(("groups", "value"), GROUP_CASES)
def test_groups_read_back_from_those_same_bytes(groups: list[QueryGroup], value: str) -> None:
    assert meta.unpack_groups(value) == groups


def test_a_truncated_group_list_is_refused() -> None:
    """The count says five and the bytes hold none.  Excel's own reader
    throws on this; so does ours, rather than reporting no groups."""
    with pytest.raises(PowerQueryError):
        meta.unpack_groups(base64.b64encode((5).to_bytes(4, "little")).decode())


# --- the section --------------------------------------------------------------


def test_the_section_serializes_as_microsofts_own_serializer_serializes_it() -> None:
    document = Metadata(items=section_items())
    assert document.serialize() == METADATA_SECTION


def test_the_section_reads_back_into_the_same_items() -> None:
    document = Metadata.parse(METADATA_SECTION)
    assert [(item.item_type, item.path) for item in document.items] == [
        (item.item_type, item.path) for item in section_items()
    ]
    assert document.content == meta.EMPTY_CONTENT
    assert document.serialize() == METADATA_SECTION


def test_an_item_with_entries_is_a_query_and_one_without_them_is_a_step() -> None:
    """Excel goes by this: a formula item whose entries are empty does not
    appear in the Queries pane, and one with any entry does."""
    document = Metadata.parse(METADATA_SECTION)
    assert [item.parts[1] for item in document.queries()] == ["With Space"]
    assert [item.parts[2] for item in document.steps_of("With Space")] == ["Source"]
    assert document.query("With Space") is not None
    assert document.query("Nothing") is None


def test_a_query_is_renamed_along_with_its_steps() -> None:
    document = Metadata.parse(METADATA_SECTION)
    document.rename_query("With Space", "Renamed")
    assert [item.path for item in document.items[1:]] == [
        "Section1/Renamed",
        "Section1/Renamed/Source",
    ]


def test_dropping_a_query_takes_its_steps_with_it() -> None:
    document = Metadata.parse(METADATA_SECTION)
    document.drop_query("With Space")
    assert [item.item_type for item in document.items] == [meta.ITEM_ALL_FORMULAS]


def test_steps_are_rewritten_in_order_right_after_their_query() -> None:
    document = Metadata.parse(METADATA_SECTION)
    document.set_steps("With Space", ["Source", "Changed Type", "Filtered"])
    assert [item.path for item in document.items[1:]] == [
        "Section1/With%20Space",
        "Section1/With%20Space/Source",
        "Section1/With%20Space/Changed%20Type",
        "Section1/With%20Space/Filtered",
    ]


def test_groups_live_on_the_document_wide_item() -> None:
    document = Metadata.parse(METADATA_SECTION)
    assert document.groups() == []
    document.set_groups([QueryGroup(GROUP_A, "Staging")])
    assert document.all_formulas().get(meta.QUERY_GROUPS) is not None
    assert document.groups() == [QueryGroup(GROUP_A, "Staging")]
    document.set_groups([])
    assert document.all_formulas().get(meta.QUERY_GROUPS) is None


def test_a_section_that_is_not_one_is_refused() -> None:
    with pytest.raises(PowerQueryError, match="at least eight bytes"):
        Metadata.parse(b"\x00\x00")
    with pytest.raises(PowerQueryError, match="does not parse"):
        Metadata.parse(b"\x00\x00\x00\x00\x04\x00\x00\x00" + b"junk" + b"\x00\x00\x00\x00")
