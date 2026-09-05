"""Refresh control on a loaded query.

Where each setting is written was measured by toggling it in Excel and
diffing the file, and the values here are the ones Excel wrote.  The live
gate closes the loop the other way: it sets them from Python and has
Excel's object model read them back.
"""

from __future__ import annotations

import re
import shutil
from pathlib import Path
from xml.etree import ElementTree

import pytest

from pyopenvba.exceptions import PowerQueryError
from pyopenvba.powerquery import PowerQueryWorkbook, RefreshSettings

FIXTURES = Path(__file__).parent / "fixtures" / "power_query"


@pytest.fixture
def loaded(tmp_path: Path) -> PowerQueryWorkbook:
    out = tmp_path / "loaded.xlsx"
    shutil.copyfile(FIXTURES / "loaded_to_sheet.xlsx", out)
    return PowerQueryWorkbook(out)


def connection(book: PowerQueryWorkbook) -> str:
    return book._opc.read("xl/connections.xml").decode("utf-8")  # pyright: ignore[reportPrivateUsage]


def query_table(book: PowerQueryWorkbook) -> str:
    return book._opc.read("xl/queryTables/queryTable1.xml").decode("utf-8")  # pyright: ignore[reportPrivateUsage]


def well_formed(book: PowerQueryWorkbook) -> None:
    for name in ("xl/connections.xml", "xl/queryTables/queryTable1.xml"):
        ElementTree.fromstring(book._opc.read(name))  # pyright: ignore[reportPrivateUsage]


# --- reading ------------------------------------------------------------------


def test_a_loaded_query_starts_with_what_excel_gave_it(loaded: PowerQueryWorkbook) -> None:
    settings = loaded.query("Loaded").refresh
    assert settings.background is True
    assert settings.interval_minutes is None
    assert settings.on_open is False
    assert settings.keep_data is True
    assert settings.in_refresh_all is True
    assert settings.enabled is True


def test_a_query_that_loads_nowhere_has_no_refresh_control(tmp_path: Path) -> None:
    out = tmp_path / "three.xlsx"
    shutil.copyfile(FIXTURES / "three_queries.xlsx", out)
    with pytest.raises(PowerQueryError, match="loads nowhere"):
        _ = PowerQueryWorkbook(out).query("Numbers").refresh


def test_the_settings_say_what_they_are(loaded: PowerQueryWorkbook) -> None:
    assert "background=True" in repr(loaded.query("Loaded").refresh)


# --- writing ------------------------------------------------------------------


def test_background_refresh_is_written_on_both_parts(loaded: PowerQueryWorkbook) -> None:
    """Excel writes the connection attribute and mirrors it on the query
    table, so both are written here."""
    loaded.query("Loaded").refresh.background = False
    assert 'background="1"' not in connection(loaded)
    assert 'backgroundRefresh="0"' in query_table(loaded)
    assert loaded.query("Loaded").refresh.background is False

    loaded.query("Loaded").refresh.background = True
    assert 'background="1"' in connection(loaded)
    assert "backgroundRefresh" not in query_table(loaded)
    well_formed(loaded)


def test_a_refresh_interval_is_minutes_on_the_connection(loaded: PowerQueryWorkbook) -> None:
    loaded.query("Loaded").refresh.interval_minutes = 45
    assert 'interval="45"' in connection(loaded)
    assert loaded.query("Loaded").refresh.interval_minutes == 45

    loaded.query("Loaded").refresh.interval_minutes = None
    assert "interval=" not in connection(loaded)
    assert loaded.query("Loaded").refresh.interval_minutes is None


def test_an_interval_that_is_not_a_number_of_minutes_is_refused(loaded: PowerQueryWorkbook) -> None:
    with pytest.raises(PowerQueryError, match="positive number of minutes"):
        loaded.query("Loaded").refresh.interval_minutes = 0


def test_refresh_on_open_is_written_on_both_parts(loaded: PowerQueryWorkbook) -> None:
    loaded.query("Loaded").refresh.on_open = True
    assert 'refreshOnLoad="1"' in connection(loaded)
    assert 'refreshOnLoad="1"' in query_table(loaded)
    assert loaded.query("Loaded").refresh.on_open is True

    loaded.query("Loaded").refresh.on_open = False
    assert "refreshOnLoad" not in connection(loaded)
    assert "refreshOnLoad" not in query_table(loaded)


def test_removing_the_data_before_saving_is_a_query_table_flag(loaded: PowerQueryWorkbook) -> None:
    """The dialog's box is the other way round, and what Excel reads is
    the query table's flag."""
    loaded.query("Loaded").refresh.keep_data = False
    assert 'removeDataOnSave="1"' in query_table(loaded)
    assert 'saveData="1"' not in connection(loaded)
    assert loaded.query("Loaded").refresh.keep_data is False

    loaded.query("Loaded").refresh.keep_data = True
    assert "removeDataOnSave" not in query_table(loaded)
    assert 'saveData="1"' in connection(loaded)


def test_leaving_refresh_all_is_an_extension_on_the_connection(loaded: PowerQueryWorkbook) -> None:
    loaded.query("Loaded").refresh.in_refresh_all = False
    raw = connection(loaded)
    assert "{DE250136-89BD-433C-8126-D09CA5730AF9}" in raw
    assert 'excludeFromRefreshAll="1"' in raw
    assert loaded.query("Loaded").refresh.in_refresh_all is False
    well_formed(loaded)

    loaded.query("Loaded").refresh.in_refresh_all = True
    assert "excludeFromRefreshAll" not in connection(loaded)
    assert "<extLst>" not in connection(loaded)
    assert loaded.query("Loaded").refresh.in_refresh_all is True
    well_formed(loaded)


def test_turning_refresh_off_is_a_query_table_flag(loaded: PowerQueryWorkbook) -> None:
    loaded.query("Loaded").refresh.enabled = False
    assert 'disableRefresh="1"' in query_table(loaded)
    assert loaded.query("Loaded").refresh.enabled is False

    loaded.query("Loaded").refresh.enabled = True
    assert "disableRefresh" not in query_table(loaded)


def test_every_setting_survives_a_save(loaded: PowerQueryWorkbook) -> None:
    settings = loaded.query("Loaded").refresh
    settings.background = False
    settings.interval_minutes = 30
    settings.on_open = True
    settings.keep_data = False
    settings.in_refresh_all = False
    settings.enabled = False
    out = loaded.save()

    again = PowerQueryWorkbook(out).query("Loaded").refresh
    assert (again.background, again.interval_minutes, again.on_open) == (False, 30, True)
    assert (again.keep_data, again.in_refresh_all, again.enabled) == (False, False, False)


def test_settings_are_read_and_written_per_query(tmp_path: Path) -> None:
    """Two loaded queries have a connection each, and one must not carry
    the other's settings."""
    out = tmp_path / "two.xlsx"
    shutil.copyfile(FIXTURES / "three_queries.xlsx", out)
    book = PowerQueryWorkbook(out)
    book.load_to_sheet("Numbers", ["N"], cell="H1")
    book.load_to_sheet("Doubled", ["N", "Twice"], cell="K1")
    book.query("Numbers").refresh.interval_minutes = 15
    book.query("Doubled").refresh.on_open = True

    assert book.query("Numbers").refresh.interval_minutes == 15
    assert book.query("Doubled").refresh.interval_minutes is None
    assert book.query("Doubled").refresh.on_open is True
    assert book.query("Numbers").refresh.on_open is False
    assert connection(book).count('interval="15"') == 1


def test_the_settings_class_can_be_built_on_its_own(loaded: PowerQueryWorkbook) -> None:
    settings = RefreshSettings(loaded._opc, "Loaded")  # pyright: ignore[reportPrivateUsage]
    assert settings.background is True


# --- the example --------------------------------------------------------------


def test_the_refresh_example_builds_what_it_describes(tmp_path: Path) -> None:
    import importlib.util
    from typing import Any

    spec = importlib.util.spec_from_file_location(
        "power_query_refresh", Path(__file__).parents[1] / "examples" / "power_query_refresh.py"
    )
    assert spec is not None and spec.loader is not None
    demo: Any = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(demo)

    out = demo.build(tmp_path / "refresh.xlsx")
    book = PowerQueryWorkbook(out)
    assert book.query_names() == ["Snapshot", "Hourly", "Manual"]
    assert book.query("Snapshot").refresh.on_open is True
    assert book.query("Hourly").refresh.interval_minutes == 60
    assert book.query("Hourly").refresh.background is False
    assert book.query("Manual").refresh.in_refresh_all is False
    assert book.query("Manual").refresh.keep_data is False
    assert len(demo.report(out)) == 3
    assert re.search(r"Hourly\s+on open: False\s+every 60 min", demo.report(out)[1])
