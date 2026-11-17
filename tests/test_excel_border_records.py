"""Excel's measured answers for borders and UsedRange, replayed against the model.

Each fixture holds probes measured in Excel: every case a macro, written as the probe's shared prefix and suffix around
the case's own middle (the numbers of its lines in the probe's table of lines), run on a new one-sheet workbook, and
the text Excel returned. A case that saves names its file
under {folder}; one that opens a book the probe built names it {book}, made here from the parts every book of the
probe shares and the parts of its own. The model must return the same text for every case, or, for a case marked
refused, refuse it with the error its "refused" text matches. A case marked "differs" is a known difference, kept with
Excel's answer and the reason the model answers otherwise; the test fails once the model answers it as Excel did, so the
mark comes off.
"""
from __future__ import annotations

import json
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

FIXTURES = Path(__file__).parent / "fixtures"
FILES = ["border_gaps.json", "border_neighbours.json", "border_union.json", "border_rounds.json", "used_range.json"]


def _probes() -> list[Any]:
    found: list[Any] = []
    for file in FILES:
        data = json.loads((FIXTURES / file).read_text(encoding="utf-8"))
        for probe in data["probes"]:
            found.append(pytest.param(probe, id=f"{file[:-5]}-{probe['name']}"))
    return found


@pytest.mark.parametrize("probe", _probes())
def test_the_model_answers_as_excel_did(probe: dict[str, Any], tmp_path: Path) -> None:
    module, _, _ = probe["entry"].rpartition(".")
    wrong: list[str] = []
    for at, case in enumerate(probe["cases"]):
        folder = tmp_path / str(at)
        folder.mkdir()
        source = probe["prefix"] + "\n".join(probe["lines"][line] for line in case["middle"]) + probe["suffix"]
        if "book" in case:
            book = folder / "book.xlsx"
            with zipfile.ZipFile(book, "w", zipfile.ZIP_DEFLATED) as z:
                for member, text in {**probe.get("book", {}), **probe["books"][case["book"]]}.items():
                    z.writestr(member, text)
            source = source.replace("{book}", str(book))
        source = source.replace("{folder}", str(folder))
        app = ExcelApplication(sandbox=False)
        app.add_workbook()
        app.add_module(source, name=module or "M")
        args = [arg.replace("{folder}", str(folder)) if isinstance(arg, str) else arg for arg in case["args"]]
        if "refused" in case:
            # A case the model does not answer yet refuses, its Excel answer kept for the work that will answer it.
            with pytest.raises(VBAUnsupportedError, match=case["refused"]):
                app.run(probe["entry"], *args)
            continue
        got = str(app.run(probe["entry"], *args))
        if "differs" in case:
            # A known difference, kept with Excel's answer and its reason: answering as Excel did means the record moves to
            # the cases that match.
            if got == case["excel"]:
                wrong.append(f"case {at} now answers as Excel did; take its differs mark away: {case['differs']}")
            continue
        if got != case["excel"]:
            wrong.append(f"case {at}:\n  Excel {case['excel'][:300]}\n  model {got[:300]}")
    assert not wrong, f"{len(wrong)} of {len(probe['cases'])} differ\n" + "\n".join(wrong[:5])
