"""Workbooks.Open of CSV and text files, replayed from live Excel (scripts/measure_text_files.py,
tests/fixtures/text_files.json): how a file splits into rows and fields and how each field is typed, how its bytes
are read, what the workbook and its sheet are called, the delimiter a Format leaves for the text files opened after
it, the arguments Open takes and the errors it raises; and the same currency and percent signs typed into a cell
and read by TextToColumns."""

from __future__ import annotations

import base64
import json
from pathlib import Path

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

RECORD = json.loads((Path(__file__).parent / "fixtures" / "text_files.json").read_text(encoding="utf-8"))
#: The cases the model reports unsupported, and what the report names: a file in UTF-16 big-endian, which holds
#: NULs, read by Excel as Windows-1252 with the NULs dropped.
UNSUPPORTED = {"utf16_big_endian": "NUL"}


@pytest.fixture(scope="module")
def replayed(tmp_path_factory: pytest.TempPathFactory) -> dict[str, str]:
    """Every case run in one application, in the order Excel ran them, as the delimiter a Format names carries
    from one open to the next."""
    folder = tmp_path_factory.mktemp("text_files").resolve()
    for entry in RECORD["files"].values():
        (folder / entry["file"]).write_bytes(base64.b64decode(entry["data"]))
    (folder / RECORD["failing"]["garbage_xlsx"]).write_bytes(b"hello")
    for steps in RECORD["sequences"].values():
        for file, _ in steps:
            (folder / file).write_bytes(RECORD["mixed"].encode("cp1252"))
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(RECORD["probe"], name="Probe")
    answers: dict[str, str] = {}
    for name, steps in RECORD["sequences"].items():
        answers[name] = str(app.run("Probe.Sequence", "\x02".join(f"{folder / file}\x01{how}" for file, how in steps)))
    for procedure in ("Typing", "Columns"):
        answers[procedure.lower()] = str(app.run(f"Probe.{procedure}", "\x01".join(RECORD["signed"])))
    planned = [(name, entry["file"]) for name, entry in RECORD["files"].items()]
    planned += [*RECORD["made"].items(), *RECORD["failing"].items()]
    for name, file in planned:
        path = str(folder / file)
        try:
            answers[name] = str(app.run("Probe.Probe", name, path)).replace(path, "{path}")
        except VBAUnsupportedError as problem:
            answers[name] = f"unsupported: {problem}"
    return answers


@pytest.mark.parametrize("name", list(RECORD["answers"]))
def test_each_open_is_excels(replayed: dict[str, str], name: str) -> None:
    if name in UNSUPPORTED:
        assert replayed[name].startswith("unsupported") and UNSUPPORTED[name] in replayed[name]
        return
    assert replayed[name] == RECORD["answers"][name]


def _pieces(answer: str) -> list[str]:
    return answer.replace("~", "|").split("|")


@pytest.mark.parametrize("name", ["typing", "columns", "txt_delimiter"])
def test_each_piece_of_the_long_answers_is_excels(replayed: dict[str, str], name: str) -> None:
    """The answers that hold many cases, piece by piece, so a difference names its case."""
    assert _pieces(replayed[name]) == _pieces(RECORD["answers"][name])
