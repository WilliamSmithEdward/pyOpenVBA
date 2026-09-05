"""The M section document a workbook keeps its queries in.

The layout expected here -- CRLF, a blank line between members, no
newline after the last one -- is Excel's own, read back from workbooks it
wrote.  The step lists are checked against the step items Excel records
in the metadata of ``tests/fixtures/power_query/step_shapes.xlsx``, which
is the same enumeration seen from the other side.
"""

from __future__ import annotations

import pytest

from pyopenvba.exceptions import PowerQueryError
from pyopenvba.powerquery._section import (
    Section,
    is_function_expression,
    let_steps,
    new_section,
    quote_name,
    rename_references,
    unquote_name,
)

CRLF = "\r\n"
SIMPLE = (
    "section Section1;" + CRLF * 2 +
    "shared Numbers = let" + CRLF +
    "    Source = {1..10}," + CRLF +
    "    Table = Table.FromList(Source)" + CRLF +
    "in" + CRLF +
    "    Table;" + CRLF * 2 +
    'shared #"Count Of Rows" = let' + CRLF +
    "    Source = Table.RowCount(Numbers)" + CRLF +
    "in" + CRLF +
    "    Source;"
)


# --- reading ------------------------------------------------------------------


def test_a_document_reads_as_its_members() -> None:
    section = Section(SIMPLE)
    assert section.name == "Section1"
    assert section.names() == ["Numbers", "Count Of Rows"]
    assert section.formula("Numbers").startswith("let")
    assert section.formula("Count Of Rows").endswith("Source")


def test_a_document_writes_back_exactly_as_it_was_read() -> None:
    assert Section(SIMPLE).text == SIMPLE


@pytest.mark.parametrize(
    ("name", "spelled"),
    [
        ("Plain", "Plain"),
        ("With Space", '#"With Space"'),
        ("Café", '#"Café"'),
        ("let", '#"let"'),
        ("9Leading", '#"9Leading"'),
        ("Quote\"Inside", '#"Quote""Inside"'),
        ("With.Dot", '#"With.Dot"'),
    ],
)
def test_a_name_is_quoted_when_it_has_to_be(name: str, spelled: str) -> None:
    assert quote_name(name) == spelled
    assert unquote_name(spelled) == name


def test_an_attribute_record_carries_the_description() -> None:
    text = (
        "section Section1;" + CRLF * 2 +
        '[ Description = "A described query" ]' + CRLF +
        "shared Described = let Source = 1 in Source;"
    )
    assert Section(text).description("Described") == "A described query"
    assert Section(SIMPLE).description("Numbers") is None


def test_a_document_that_is_not_one_is_refused() -> None:
    with pytest.raises(PowerQueryError, match="does not start with 'section'"):
        Section("shared X = 1;")
    with pytest.raises(PowerQueryError, match="no ';'"):
        Section("section Section1")


def test_an_unfinished_literal_is_refused() -> None:
    with pytest.raises(PowerQueryError, match="never closed"):
        Section('section Section1;' + CRLF + 'shared X = "open;')
    with pytest.raises(PowerQueryError, match="never closed"):
        Section("section Section1;" + CRLF + "shared X = 1; /* open")


# --- steps --------------------------------------------------------------------


@pytest.mark.parametrize(
    ("formula", "steps"),
    [
        ("let A = 1, B = A, C = B in C", ["A", "B", "C"]),
        ('let #"Step One" = 1, #"Changed Type" = 2 in 3', ["Step One", "Changed Type"]),
        ("let Outer = let Inner = 1, Inner2 = Inner in Inner2, Final = Outer in Final", ["Outer", "Final"]),
        ("1 + 1", []),
        ("(n as number) => let Doubled = n * 2 in Doubled", []),
        ('let Rec = [a = 1, b = 2], Lst = {1, 2, 3}, Txt = "a, b, c" in Txt', ["Rec", "Lst", "Txt"]),
        ("let F = each _ + 1, G = F(1) in G", ["F", "G"]),
        ("let // a comment with ; and ,\n    A = 1, /* block , ; */ B = A in B", ["A", "B"]),
        ("   let A = 1 in A", ["A"]),
        ("letter", []),
    ],
)
def test_steps_are_the_top_level_let_bindings(formula: str, steps: list[str]) -> None:
    assert let_steps(formula) == steps


@pytest.mark.parametrize(
    ("formula", "is_function"),
    [
        ("(x) => x + 1", True),
        ("(x as number) as number => x + 1", True),
        ("each _ + 1", True),
        ("()  =>  1", True),
        ("let A = 1 in A", False),
        ("[a = 1]", False),
        ("(1 + 2) * 3", False),
    ],
)
def test_a_function_is_recognised_by_how_it_is_written(formula: str, is_function: bool) -> None:
    assert is_function_expression(formula) is is_function


# --- editing ------------------------------------------------------------------


def test_a_formula_change_leaves_every_other_byte_alone() -> None:
    section = Section(SIMPLE)
    section.set_formula("Numbers", "let Source = {1..25} in Source")
    assert section.formula("Numbers") == "let Source = {1..25} in Source"
    assert section.names() == ["Numbers", "Count Of Rows"]
    assert section.formula("Count Of Rows") == Section(SIMPLE).formula("Count Of Rows")
    assert section.text.startswith("section Section1;" + CRLF * 2 + "shared Numbers = let Source")


def test_a_member_is_added_the_way_excel_adds_one() -> None:
    section = Section(SIMPLE)
    section.add("Extra", "let" + CRLF + "    Source = 99" + CRLF + "in" + CRLF + "    Source")
    assert section.names() == ["Numbers", "Count Of Rows", "Extra"]
    assert section.text.endswith(
        CRLF * 2 + "shared Extra = let" + CRLF + "    Source = 99" + CRLF + "in" + CRLF + "    Source;"
    )
    assert not section.text.endswith(CRLF)


def test_a_member_added_with_a_description_carries_its_record() -> None:
    section = new_section()
    section.add("Q", "1", "why it exists")
    assert section.text == 'section Section1;\r\n\r\n[ Description = "why it exists" ]\r\nshared Q = 1;'
    assert section.description("Q") == "why it exists"


def test_a_description_can_be_added_changed_and_dropped() -> None:
    section = Section(SIMPLE)
    section.set_description("Numbers", "one to ten")
    assert section.description("Numbers") == "one to ten"
    section.set_description("Numbers", 'a "quoted" one')
    assert section.description("Numbers") == 'a "quoted" one'
    section.set_description("Numbers", None)
    assert section.description("Numbers") is None
    assert section.text == SIMPLE


def test_removing_a_member_takes_its_blank_line_with_it() -> None:
    section = Section(SIMPLE)
    section.remove("Numbers")
    assert section.names() == ["Count Of Rows"]
    assert section.text == (
        "section Section1;" + CRLF * 2 + 'shared #"Count Of Rows" = let' + CRLF +
        "    Source = Table.RowCount(Numbers)" + CRLF + "in" + CRLF + "    Source;"
    )


def test_removing_the_last_member_leaves_no_trailing_blank_line() -> None:
    section = Section(SIMPLE)
    section.remove("Count Of Rows")
    assert section.text.endswith("    Table;")


def test_a_rename_can_carry_the_references_with_it() -> None:
    section = Section(SIMPLE)
    section.rename("Numbers", "Renamed Numbers", update_references=True)
    assert section.names() == ["Renamed Numbers", "Count Of Rows"]
    assert '#"Renamed Numbers"' in section.formula("Count Of Rows")
    assert "Numbers" not in section.formula("Count Of Rows").replace('#"Renamed Numbers"', "")


def test_a_rename_can_leave_the_references_alone() -> None:
    section = Section(SIMPLE)
    section.rename("Numbers", "Renamed", update_references=False)
    assert "Table.RowCount(Numbers)" in section.formula("Count Of Rows")


def test_a_rename_onto_a_name_already_there_is_refused() -> None:
    with pytest.raises(PowerQueryError, match="already has a member"):
        Section(SIMPLE).rename("Numbers", "Count Of Rows")


def test_adding_a_name_already_there_is_refused() -> None:
    with pytest.raises(PowerQueryError, match="already has a member"):
        Section(SIMPLE).add("Numbers", "1")


def test_asking_for_a_member_that_is_not_there_is_refused() -> None:
    with pytest.raises(PowerQueryError, match="no member named"):
        Section(SIMPLE).formula("Nothing")


# --- references ---------------------------------------------------------------


@pytest.mark.parametrize(
    ("text", "renamed"),
    [
        ("Numbers", '#"N B"'),
        ("Table.RowCount(Numbers)", 'Table.RowCount(#"N B")'),
        ('"Numbers"', '"Numbers"'),
        ("// Numbers", "// Numbers"),
        ("/* Numbers */", "/* Numbers */"),
        ("[Numbers]", "[Numbers]"),
        ("[Numbers = 1]", "[Numbers = 1]"),
        ("[a = Numbers]", '[a = #"N B"]'),
        ("[a = 1, Numbers = 2]", "[a = 1, Numbers = 2]"),
        ("NumbersTwo", "NumbersTwo"),
        ('#"Numbers"', '#"N B"'),
        ("each [Numbers]", "each [Numbers]"),
        ("{Numbers, 1}", '{#"N B", 1}'),
    ],
)
def test_a_reference_is_renamed_only_where_it_is_a_reference(text: str, renamed: str) -> None:
    assert rename_references(text, "Numbers", "N B") == renamed


# --- building from nothing ----------------------------------------------------


def test_an_empty_document_is_the_one_excel_starts_from() -> None:
    assert new_section().text == "section Section1;"
    assert new_section().names() == []
