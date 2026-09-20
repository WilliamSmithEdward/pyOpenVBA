"""The M evaluator: the language, the library, and what it refuses.

The semantics that can be checked against live Excel are in
``test_powerquery_refresh.py``; this covers the shape of the language
itself and the parts of the library a query leans on.
"""

from __future__ import annotations

import datetime as _dt

import pytest

from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.mlang import MError, MSyntaxError, Record, Table, evaluate_query


def run(source: str) -> object:
    return evaluate_query(source)


def rows(source: str) -> list[list[object]]:
    value = run(source)
    assert isinstance(value, Table)
    return [list(value.columns), *[list(row) for row in value.rows]]


# --- the language --------------------------------------------------------------------


def test_arithmetic_and_precedence() -> None:
    assert run("1 + 2 * 3") == 7
    assert run("(1 + 2) * 3") == 9


def test_text_is_joined_with_the_ampersand() -> None:
    assert run('"a" & "b"') == "ab"


def test_a_let_binds_its_steps() -> None:
    assert run("let x = 1, y = x + 1 in y") == 2


def test_a_let_step_may_name_one_written_after_it() -> None:
    """Bindings are lazy, so order in the source is not order of work."""
    assert run("let y = x + 1, x = 1 in y") == 2


def test_a_cycle_between_steps_is_reported() -> None:
    with pytest.raises(MError) as raised:
        run("let a = b, b = a in a")
    assert "cyclic" in str(raised.value).lower()


def test_a_record_field_may_name_another() -> None:
    value = run("[a = 1, b = a + 1]")
    assert isinstance(value, Record)
    assert value.fields == {"a": 1, "b": 2}


def test_a_bracket_is_a_field_of_the_row_inside_an_each() -> None:
    assert rows('Table.SelectRows(#table({"A"}, {{1}, {2}, {3}}), each [A] > 1)') == [
        ["A"],
        [2],
        [3],
    ]


def test_each_underscore_is_the_argument() -> None:
    assert run("List.Transform({1, 2}, each _ * 10)") == [10, 20]


def test_a_function_can_be_written_and_called() -> None:
    assert run("let f = (x) => x * 2 in f(21)") == 42


def test_an_optional_argument_may_be_left_out() -> None:
    assert run("let f = (x, optional y) => x + (if y = null then 0 else y) in f(1)") == 1


def test_if_takes_one_branch() -> None:
    assert run('if 1 > 2 then "y" else "n"') == "n"


def test_try_otherwise_catches_a_raised_error() -> None:
    assert run('try error "boom" otherwise "caught"') == "caught"


def test_try_without_otherwise_gives_the_error_record() -> None:
    value = run('try error "boom"')
    assert isinstance(value, Record)
    assert value.get("HasError") is True


def test_coalesce_takes_the_first_that_is_not_null() -> None:
    assert run("null ?? 5") == 5


def test_a_quoted_name_may_have_spaces() -> None:
    assert run('let #"a b" = 2 in #"a b" * 3') == 6


def test_text_escapes_are_read() -> None:
    assert run('"a#(lf)b"') == "a\nb"


def test_a_date_literal_is_a_date() -> None:
    assert run("#date(2021, 3, 4)") == _dt.date(2021, 3, 4)


def test_source_that_is_not_m_is_a_syntax_error() -> None:
    with pytest.raises(MSyntaxError):
        run("let x = in x")


# --- the library ----------------------------------------------------------------------


def test_a_literal_table_carries_its_columns() -> None:
    assert rows('#table({"A", "B"}, {{1, 2}, {3, 4}})') == [["A", "B"], [1, 2], [3, 4]]


def test_add_column_sees_the_row() -> None:
    assert rows('Table.AddColumn(#table({"A"}, {{1}, {2}}), "Twice", each [A] * 2)') == [
        ["A", "Twice"],
        [1, 2],
        [2, 4],
    ]


def test_group_aggregates_each_group() -> None:
    assert rows(
        'Table.Group(#table({"K", "V"}, {{"a", 1}, {"a", 2}, {"b", 5}}), {"K"},'
        ' {{"Total", each List.Sum([V])}})'
    ) == [["K", "Total"], ["a", 3], ["b", 5]]


def test_sort_takes_a_direction() -> None:
    assert rows('Table.Sort(#table({"A"}, {{3}, {1}, {2}}), {{"A", Order.Descending}})') == [
        ["A"],
        [3],
        [2],
        [1],
    ]


def test_a_join_brings_the_other_table_in() -> None:
    assert rows(
        'Table.ExpandTableColumn('
        'Table.NestedJoin(#table({"K"}, {{1}, {2}}), {"K"},'
        ' #table({"K", "V"}, {{1, "x"}}), {"K"}, "N"), "N", {"V"})'
    ) == [["K", "V"], [1, "x"], [2, None]]


def test_promote_headers_takes_the_first_row() -> None:
    assert rows('Table.PromoteHeaders(#table({"C1", "C2"}, {{"A", "B"}, {1, 2}}))') == [
        ["A", "B"],
        [1, 2],
    ]


def test_transform_column_types_converts() -> None:
    assert rows(
        'Table.TransformColumnTypes(#table({"A"}, {{"12"}}), {{"A", Int64.Type}})'
    ) == [["A"], [12]]


def test_csv_is_read_with_its_quoting() -> None:
    assert rows('Csv.Document("a,b#(lf)1,""x,y""")') == [
        ["Column1", "Column2"],
        ["a", "b"],
        ["1", "x,y"],
    ]


def test_json_becomes_records_and_lists() -> None:
    value = run('Json.Document("{""x"": [1, 2]}")')
    assert isinstance(value, Record)
    assert value.get("x") == [1, 2]


def test_list_functions_cover_the_common_ground() -> None:
    assert run("List.Sum({1, 2, 3})") == 6
    assert run("List.Average({1, 2, 3})") == 2
    assert run("List.Max({1, 5, 3})") == 5
    assert run('List.Distinct({1, 1, 2})') == [1, 2]
    assert run("List.Select({1, 2, 3}, each _ > 1)") == [2, 3]


def test_text_functions_cover_the_common_ground() -> None:
    assert run('Text.Upper("ab")') == "AB"
    assert run('Text.Split("a,b", ",")') == ["a", "b"]
    assert run('Text.Combine({"a", "b"}, "-")') == "a-b"
    assert run('Text.Start("abc", 2)') == "ab"
    assert run('Text.End("abc", 5)') == "abc"


# --- what it refuses ---------------------------------------------------------------------


def test_a_real_function_this_lacks_says_so() -> None:
    with pytest.raises(VBAUnsupportedError) as raised:
        run('Table.Profile(#table({"A"}, {{1}}))')
    assert "Table.Profile" in str(raised.value)


def test_a_source_this_cannot_reach_says_so() -> None:
    with pytest.raises(VBAUnsupportedError) as raised:
        run('Sql.Database("server", "db")')
    assert "Sql.Database" in str(raised.value)


def test_a_name_m_has_never_had_is_not_recognised() -> None:
    with pytest.raises(MError) as raised:
        run("NoSuchName")
    assert "wasn't recognized" in str(raised.value)
