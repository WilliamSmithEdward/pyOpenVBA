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


def normal(text: str, obj: str) -> str:
    """Application's settings start as the measuring harness left them, so only what the set did is compared there."""
    text = re.sub(r"Book\d+", "Book", text)
    return text.split("|", 1)[1] if obj == "Application" and "|" in text else text


def run_case(source: str, which: int) -> str:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(source, name="Measure")
    return str(app.run("T", which))


def compare() -> list[tuple[str, str, str, str, str]]:
    """Each case: its property, its value, the verdict, Excel's answer and the model's."""
    cases: list[tuple[str, str, str, str, str]] = []
    rows = json.loads(FIXTURE.read_text(encoding="utf-8"))["rows"]
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
            verdict = "agrees" if normal(model, row["object"]) == normal(answer, row["object"]) else "differs"
            cases.append((name, value, verdict, answer, model))
    return cases


def main() -> None:
    cases = compare()
    totals: dict[str, Counter[str]] = {}
    for name, value, verdict, _, _ in cases:
        totals.setdefault(f"{name.split('.')[0]} {value.split()[0]}", Counter())[verdict] += 1
    for obj, count in totals.items():
        total = sum(count.values())
        measured = total - count["unmeasured"]
        print(f"{obj}: {total} cases, {count['agrees']} agree, {count['differs']} differ, {count['reports']} report, "
              f"{count['unmeasured']} unmeasured; disagreement {100 * count['differs'] / max(measured, 1):.1f}%")
    if "--show" in sys.argv:
        for name, value, verdict, answer, model in cases:
            if verdict in ("differs", "reports", "unmeasured"):
                print(f"{verdict[:6]} {name} = {value}: excel {answer} | model {model}")


if __name__ == "__main__":
    main()
