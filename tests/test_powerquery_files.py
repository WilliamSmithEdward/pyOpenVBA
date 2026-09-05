"""Queries as files on disk, and the command line over them."""

from __future__ import annotations

import json
import shutil
from pathlib import Path

import pytest

from pyopenvba import pull_power_query, push_power_query
from pyopenvba.__main__ import main
from pyopenvba.exceptions import PowerQueryError
from pyopenvba.powerquery import PowerQueryWorkbook
from pyopenvba.powerquery._files import MANIFEST, file_name

FIXTURES = Path(__file__).parent / "fixtures" / "power_query"


@pytest.fixture
def workbook(tmp_path: Path) -> Path:
    out = tmp_path / "three.xlsx"
    shutil.copyfile(FIXTURES / "three_queries.xlsx", out)
    return out


def test_every_query_becomes_a_file_beside_a_manifest(workbook: Path, tmp_path: Path) -> None:
    written = pull_power_query(workbook, tmp_path / "src")
    names = {path.name for path in written}
    assert names == {"Numbers.m", "Doubled.m", "Count Of Rows.m", MANIFEST}
    assert (tmp_path / "src" / "Numbers.m").read_text(encoding="utf-8").startswith("let")
    manifest = json.loads((tmp_path / "src" / MANIFEST).read_text(encoding="utf-8"))
    assert [record["name"] for record in manifest["queries"]] == ["Numbers", "Doubled", "Count Of Rows"]


def test_a_pull_and_a_push_leave_the_queries_as_they_were(workbook: Path, tmp_path: Path) -> None:
    before = PowerQueryWorkbook(workbook)
    formulas = {name: before.query(name).formula for name in before.query_names()}
    pull_power_query(workbook, tmp_path / "src")
    push_power_query(tmp_path / "src", workbook)
    after = PowerQueryWorkbook(workbook)
    assert {name: after.query(name).formula for name in after.query_names()} == formulas


def test_an_edited_file_reaches_the_workbook(workbook: Path, tmp_path: Path) -> None:
    pull_power_query(workbook, tmp_path / "src")
    (tmp_path / "src" / "Numbers.m").write_text("let Source = {1..3} in Source", encoding="utf-8")
    touched = push_power_query(tmp_path / "src", workbook)
    assert "Numbers" in touched
    assert PowerQueryWorkbook(workbook).query("Numbers").formula == "let Source = {1..3} in Source"


def test_a_new_file_becomes_a_new_query(workbook: Path, tmp_path: Path) -> None:
    pull_power_query(workbook, tmp_path / "src")
    (tmp_path / "src" / "Fresh.m").write_text("let Source = 1 in Source", encoding="utf-8")
    push_power_query(tmp_path / "src", workbook)
    book = PowerQueryWorkbook(workbook)
    assert "Fresh" in book.query_names()
    assert book.query("Fresh").steps == ["Source"]


def test_a_query_the_directory_does_not_hold_stays_unless_asked(workbook: Path, tmp_path: Path) -> None:
    pull_power_query(workbook, tmp_path / "src")
    (tmp_path / "src" / "Doubled.m").unlink()
    push_power_query(tmp_path / "src", workbook)
    assert "Doubled" in PowerQueryWorkbook(workbook).query_names()
    push_power_query(tmp_path / "src", workbook, remove_missing=True)
    assert "Doubled" not in PowerQueryWorkbook(workbook).query_names()


def test_a_name_a_file_system_would_refuse_still_round_trips(workbook: Path, tmp_path: Path) -> None:
    """The manifest is what carries the real name."""
    book = PowerQueryWorkbook(workbook)
    book.add_query("Sales/EU", "let Source = 1 in Source")
    book.save()
    pull_power_query(workbook, tmp_path / "src")
    assert (tmp_path / "src" / "Sales_EU.m").is_file()
    (tmp_path / "src" / "Sales_EU.m").write_text("let Source = 2 in Source", encoding="utf-8")
    push_power_query(tmp_path / "src", workbook)
    assert PowerQueryWorkbook(workbook).query("Sales/EU").formula == "let Source = 2 in Source"


def test_two_names_that_map_to_one_file_get_two_files() -> None:
    taken: set[str] = set()
    assert file_name("a/b", taken) == "a_b.m"
    assert file_name("a:b", taken) == "a_b_2.m"


def test_a_reserved_device_name_is_kept_out_of_the_way() -> None:
    assert file_name("CON", set()) == "_CON.m"


def test_pushing_from_somewhere_that_is_not_a_directory_is_refused(workbook: Path, tmp_path: Path) -> None:
    with pytest.raises(PowerQueryError, match="not a directory"):
        push_power_query(tmp_path / "nowhere", workbook)


def test_a_manifest_that_does_not_parse_is_refused(workbook: Path, tmp_path: Path) -> None:
    pull_power_query(workbook, tmp_path / "src")
    (tmp_path / "src" / MANIFEST).write_text("{", encoding="utf-8")
    with pytest.raises(PowerQueryError, match="does not parse"):
        push_power_query(tmp_path / "src", workbook)


# --- the command line ---------------------------------------------------------


def test_the_command_line_lists_the_queries(capsys: pytest.CaptureFixture[str]) -> None:
    assert main(["pq-ls", str(FIXTURES / "grouped.xlsx")]) == 0
    out = capsys.readouterr().out
    assert "Numbers\tconnection-only" in out
    assert "[Staging]" in out


def test_the_command_line_pulls_and_pushes(workbook: Path, tmp_path: Path, capsys: pytest.CaptureFixture[str]) -> None:
    assert main(["pq-pull", str(workbook), str(tmp_path / "src")]) == 0
    assert (tmp_path / "src" / "Numbers.m").is_file()
    capsys.readouterr()
    (tmp_path / "src" / "Numbers.m").write_text("let Source = 7 in Source", encoding="utf-8")
    assert main(["pq-push", str(tmp_path / "src"), str(workbook)]) == 0
    assert "Numbers" in capsys.readouterr().out
    assert PowerQueryWorkbook(workbook).query("Numbers").formula == "let Source = 7 in Source"


def test_the_command_line_can_push_to_a_new_file(workbook: Path, tmp_path: Path) -> None:
    main(["pq-pull", str(workbook), str(tmp_path / "src")])
    (tmp_path / "src" / "Numbers.m").write_text("let Source = 8 in Source", encoding="utf-8")
    out = tmp_path / "copy.xlsx"
    assert main(["pq-push", str(tmp_path / "src"), str(workbook), "--out", str(out)]) == 0
    assert PowerQueryWorkbook(out).query("Numbers").formula == "let Source = 8 in Source"
    assert PowerQueryWorkbook(workbook).query("Numbers").formula.startswith("let\r\n    Source = {1..10}")


# --- the example ---------------------------------------------------------------


def test_the_demo_builds_the_workbook_it_describes(tmp_path: Path) -> None:
    """examples/power_query_demo.py is part of the documentation, so it
    has to keep working."""
    import importlib.util

    spec = importlib.util.spec_from_file_location(
        "power_query_demo", Path(__file__).parents[1] / "examples" / "power_query_demo.py"
    )
    assert spec is not None and spec.loader is not None
    demo = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(demo)

    out = demo.build(tmp_path / "demo.xlsx")
    book = PowerQueryWorkbook(out)
    assert [group.name for group in book.groups()] == ["Parameters", "Functions", "Web", "Workbook"]
    assert len(book.query_names()) == 12
    assert book.query("GetFromPokeApi").is_function
    assert book.query("TitleCase").is_function
    assert book.query("Pokedex").group is not None
    assert book.query("Pokedex").group.name == "Web"  # pyright: ignore[reportOptionalMemberAccess]
    loaded = [query.name for query in book.queries() if query.load_target == "table"]
    assert loaded == ["Pokedex", "PokemonStats", "Earthquakes", "Rates", "RateHistory", "OrderSummary"]
    assert all(query.description for query in book.queries())
    assert "pokeapi.co" in book.query("GetFromPokeApi").formula
