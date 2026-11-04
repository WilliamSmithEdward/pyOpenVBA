"""Application.Path, a chosen install path rather than a measured one, and LibraryPath's measured relation to it
(scripts/measure_application_paths.py)."""

from __future__ import annotations

import json
from pathlib import Path

import pytest

from pyopenvba.apps.excel import ExcelApplication

MEASURED = json.loads((Path(__file__).parent / "fixtures/application_paths.json").read_text(encoding="utf-8"))
RELATIONS = dict(part.split("=", 1) for part in MEASURED["relations"].split("|"))
INSTALL = "C:\\Program Files\\Microsoft Office\\root\\Office16"


def _read(expression: str) -> object:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(f"Function Answer() As Variant\nAnswer = {expression}\nEnd Function\n", name="M")
    return app.run("Answer")


def test_path_is_the_chosen_install_folder() -> None:
    assert _read("Application.Path") == INSTALL
    assert RELATIONS["office16"] == "True"


def test_path_does_not_follow_the_current_folder(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.chdir(tmp_path)
    assert _read("Application.Path") == INSTALL


def test_library_path_stands_under_path_as_excels_does() -> None:
    under, tail = RELATIONS["library"].split(":")
    assert under == "True"
    assert _read("Application.LibraryPath") == INSTALL + tail


def test_the_folders_of_the_user_are_not_under_path() -> None:
    for name in ("startup", "templates", "user_library", "default_file"):
        assert RELATIONS[name] == "False"
