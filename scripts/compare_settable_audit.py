"""The disagreement rate of the model against tests/fixtures/settable_audit.json, per object: each case agrees, differs, or
the model reports it as not implemented.

    python scripts/compare_settable_audit.py [--show]
"""
from __future__ import annotations

import json
import re
import sys
from collections import Counter
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from pyopenvba.apps.excel import ExcelApplication  # noqa: E402
from pyopenvba.exceptions import VBAUnsupportedError  # noqa: E402

FIXTURE = ROOT / "tests/fixtures/settable_audit.json"


#: What a macro would sensibly set each kind of property to; anything else is a bad value. A Boolean's sensible values are
#: those VBA itself would write for one, and a text property's are ordinary text.
BOOLEAN_VALID = {"True", "False", "0", "1", "-1"}
ENUM_VALID = {
    "Worksheet.EnableSelection": {"0", "1", "-4142"}, "Worksheet.Visible": {"-1", "0", "2", "xlSheetVeryHidden", "xlSheetHidden"},
    "Workbook.AccuracyVersion": {"0", "1", "2"}, "Workbook.AutoUpdateFrequency": {"1", "2", "3"},
    "Workbook.ChangeHistoryDuration": {"1", "2", "3"}, "Workbook.CompatibilityVersion": {"1", "2"},
    "Workbook.ConflictResolution": {"1", "2", "3"}, "Workbook.DisplayDrawingObjects": {"-4104", "2", "3"},
    "Workbook.ExternalCodeServiceTimeout": {"1", "2", "3"}, "Workbook.UpdateLinks": {"1", "2", "3"},
    "Application.Calculation": {"-4105", "-4135", "2"},
    "Application.Cursor": {"-4143", "1", "2", "3", "xlWait", "xlIBeam", "xlNorthwestArrow"},
    "Application.CutCopyMode": {"False", "xlCopy", "xlCut", "1", "2", "0"},
    "Worksheet.StandardWidth": {"1", "2", "3", "1.5", "255", "8.43"},
    "Worksheet.Name": {'"abc"', '"Other"', '"Sheet9"'},
    "Worksheet.ScrollArea": {'""', '"A1:B2"', '"$A$1"'},
    "Application.StatusBar": {'"abc"', '""', "False", '"x"', '"True"'},
}
TEXT_VALID = {'"abc"', '""', '"x"'}
#: Properties whose values are names of styles; none of the values tried is one, so all count as bad.
STYLES = {"Workbook.DefaultPivotTableStyle", "Workbook.DefaultSlicerStyle", "Workbook.DefaultTableStyle",
          "Workbook.DefaultTimelineStyle"}


def is_valid(name: str, value: str, tried: list[str]) -> bool:
    """Whether a macro would sensibly set the property to this value; what was tried says its kind."""
    if name in STYLES:
        return False
    if name in ENUM_VALID:
        return value in ENUM_VALID[name]
    return value in (BOOLEAN_VALID if '"True"' in tried else TEXT_VALID if '"abc"' in tried else BOOLEAN_VALID & {"0", "1"})


#: Application's settings, and DoNotPromptForConvert, which the user's registry keeps, start as the measuring machine left
#: them, so only what the set did is compared there.
FROM_THE_MACHINE = {"Application", "Workbook.DoNotPromptForConvert"}


def normal(text: str, obj: str) -> str:
    text = re.sub(r"Book\d+", "Book", text)
    return text.split("|", 1)[1] if obj in FROM_THE_MACHINE and "|" in text else text


def run_case(source: str, which: int) -> str:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(source, name="Measure")
    return str(app.run("T", which))


ROWS = json.loads(FIXTURE.read_text(encoding="utf-8"))["rows"]


def compare() -> list[tuple[str, str, str, str, str]]:
    """Each case: its property, its value, the verdict, Excel's answer and the model's."""
    cases: list[tuple[str, str, str, str, str]] = []
    rows = ROWS
    for name, row in rows.items():
        values = [*(f"late {one}" for one in row["values"]), *(f"typed {one}" for one in row["values"])]
        for index, (value, answer) in enumerate(zip(values, row["answers"])):
            if answer.startswith("failed"):
                cases.append((name, value, "unmeasured", answer, ""))
                continue
            try:
                model = run_case(row["source"], index)
            except VBAUnsupportedError as failure:
                cases.append((name, value, "reports", answer, str(failure)[:70]))
                continue
            except Exception as failure:  # noqa: BLE001 - a crash in the model is a disagreement to count and show
                cases.append((name, value, "differs", answer, f"{type(failure).__name__}: {failure}"[:90]))
                continue
            key = name if name in FROM_THE_MACHINE else row["object"]
            verdict = "agrees" if normal(model, key) == normal(answer, key) else "differs"
            cases.append((name, value, verdict, answer, model))
    return cases


def main() -> None:
    cases = compare()
    totals: dict[str, Counter[str]] = {}
    for name, value, verdict, _, _ in cases:
        binding, given = value.split(" ", 1)
        kind = "valid" if is_valid(name, given, ROWS[name]["values"]) else "bad"
        totals.setdefault(f"{name.split('.')[0]} {binding} {kind}", Counter())[verdict] += 1
    for key in sorted(totals):
        count = totals[key]
        total = sum(count.values())
        measured = total - count["unmeasured"]
        print(f"{key}: {total} cases, {count['agrees']} agree, {count['differs']} differ, {count['reports']} report, "
              f"{count['unmeasured']} unmeasured; disagreement {100 * count['differs'] / max(measured, 1):.1f}%")
    if "--show" in sys.argv:
        for name, value, verdict, answer, model in cases:
            binding, given = value.split(" ", 1)
            if verdict in ("differs", "reports", "unmeasured"):
                kind = "valid" if is_valid(name, given, ROWS[name]["values"]) else "bad"
                print(f"{verdict[:6]} {kind} {binding} {name} = {given}: excel {answer} | model {model}")


if __name__ == "__main__":
    main()
