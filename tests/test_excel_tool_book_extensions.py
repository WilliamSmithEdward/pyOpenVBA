"""Excel's 2010 extensions as another tool writes them, under other prefixes too (scripts/measure_tool_book_extensions.py). What
reads them goes by the namespace, so x14:, xm: and mc: need no respelling at open, and a save under another prefix still reads as
Excel's own."""

from __future__ import annotations

import importlib.util
import json
import sys
from pathlib import Path
from types import ModuleType
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

HERE = Path(__file__).parent / "fixtures" / "tool_books"
RECORD: dict[str, Any] = json.loads((HERE / "extensions.json").read_text(encoding="utf-8"))
READS: list[str] = RECORD["reads"]


def _script() -> ModuleType:
    folder = Path(__file__).parents[1] / "scripts"
    sys.path.insert(0, str(folder))
    spec = importlib.util.spec_from_file_location("measure_tool_book_extensions", folder / "measure_tool_book_extensions.py")
    assert spec is not None and spec.loader is not None
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


SCRIPT = _script()


def _app() -> ExcelApplication:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    return app


@pytest.mark.parametrize("variant", RECORD["variants"])
def test_reads_what_excel_reads(variant: str, tmp_path: Path) -> None:
    file = tmp_path / f"{variant}.xlsx"
    SCRIPT.made(HERE / "extensions.xlsx", variant, file)
    app = _app()
    app.add_module(SCRIPT.source(READS), name="M")
    answers = str(app.run("Go", file.as_posix())).split("\t")
    assert list(zip(READS, answers, strict=True)) == list(zip(READS, RECORD["variants"][variant]["excel"], strict=True))


@pytest.mark.parametrize("key", [key for key, entry in RECORD["edits"].items() if not entry["reported"]])
def test_excel_read_the_models_save_as_its_own(key: str, tmp_path: Path) -> None:
    """Measured live. The edit is made again so a save that starts to report or crash shows here too."""
    entry = RECORD["edits"][key]
    assert entry["excel"] == entry["model"]
    variant, _, name = key.partition("|")
    file = tmp_path / f"{variant}.xlsx"
    SCRIPT.made(HERE / "extensions.xlsx", variant, file)
    app = _app()
    app.add_module(SCRIPT.ACT.format(action=SCRIPT.EDITS[name]), name="M")
    app.run("Act", file.as_posix())
