"""Workbooks.OpenText, replayed from live Excel (scripts/measure_open_text.py, tests/fixtures/open_text.json): the
delimiters, qualifiers, start rows, field kinds, fixed widths, separators and code pages it reads with, the
settings it remembers for the next OpenText and shares with TextToColumns and Workbooks.Open, and its errors."""

from __future__ import annotations

import base64
import json
from pathlib import Path

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

RECORD = json.loads((Path(__file__).parent / "fixtures" / "open_text.json").read_text(encoding="utf-8"))
#: The cases the model reports unsupported, and what the report names: the columns Excel guesses for fixed widths,
#: a field of 32,759 characters or more, and the Macintosh code page past ASCII.
UNSUPPORTED = {"fixed_default": "guesses the columns", "default_aligned": "fixed width", "long_field": "32,759",
               "origin_mac": "Macintosh"}


@pytest.fixture(scope="module")
def replayed(tmp_path_factory: pytest.TempPathFactory) -> dict[str, str]:
    """Every case run in one application, in the order Excel ran them, as the settings one OpenText remembers carry
    to the next."""
    folder = tmp_path_factory.mktemp("open_text").resolve()
    for entry in RECORD["files"].values():
        (folder / entry["file"]).write_bytes(base64.b64decode(entry["data"]))
    for file in RECORD["sequence"]:
        (folder / file).write_bytes(RECORD["mixed"].encode("cp1252"))
    for file in RECORD["shared"]:
        (folder / file).write_bytes(base64.b64decode(RECORD["shared_text"]))
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(RECORD["probe"], name="Probe")

    def run(procedure: str, *arguments: str) -> str:
        try:
            return str(app.run(f"Probe.{procedure}", *arguments))
        except VBAUnsupportedError as problem:
            return f"unsupported: {problem}"

    answers: dict[str, str] = {}
    for name in RECORD["defaults"]:
        answers[name] = run("Probe", name, str(folder / RECORD["files"][name]["file"]))
    answers["sequence"] = run("Sequence", "\x01".join(f"{file}\x02{folder / file}" for file in RECORD["sequence"]))
    answers["shared"] = run("Remembered", "\x01".join(str(folder / file) for file in RECORD["shared"]))
    for name, entry in RECORD["files"].items():
        if name not in answers:
            answers[name] = run("Probe", name, str(folder / entry["file"]))
    missing = str(folder / "absent.txt")
    answers["missing"] = run("Failing", missing).replace(missing, "{path}")
    return answers


@pytest.mark.parametrize("name", list(RECORD["answers"]))
def test_each_open_is_excels(replayed: dict[str, str], name: str) -> None:
    if name in UNSUPPORTED:
        assert replayed[name].startswith("unsupported") and UNSUPPORTED[name] in replayed[name]
        return
    assert replayed[name].replace("~", "|").split("|") == RECORD["answers"][name].replace("~", "|").split("|")
