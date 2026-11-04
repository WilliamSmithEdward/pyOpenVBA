"""The disagreement rate of the model against tests/fixtures/settable_lists.json, per object and binding: each case agrees,
differs, or the model reports it as not implemented.

    python scripts/compare_settable_lists.py [--show]
"""
from __future__ import annotations

import json
import re
import sys
import tempfile
from collections import Counter
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / 'src'))

from pyopenvba.apps.excel import ExcelApplication  # noqa: E402
from pyopenvba.exceptions import VBAUnsupportedError  # noqa: E402

FIXTURE = ROOT / 'tests/fixtures/settable_lists.json'


def compare() -> list[tuple[str, str, str, str, str]]:
    """Each case: its property, its binding and value, the verdict, Excel's answer and the model's."""
    cases: list[tuple[str, str, str, str, str]] = []
    rows = json.loads(FIXTURE.read_text(encoding='utf-8'))['rows']
    with tempfile.TemporaryDirectory() as folder:
        for name, row in rows.items():
            values = [*(f'late {one}' for one in row['values']), *(f'typed {one}' for one in row['values'])]
            for index, (value, answer) in enumerate(zip(values, row['answers'])):
                if answer.startswith('failed'):
                    cases.append((name, value, 'unmeasured', answer, ''))
                    continue
                try:
                    app = ExcelApplication(sandbox=False)
                    app.add_workbook()
                    app.add_module(row['source'], name='Measure')
                    model = str(app.run('T', index, folder))
                except VBAUnsupportedError as failure:
                    cases.append((name, value, 'reports', answer, str(failure)[:70]))
                    continue
                except Exception as failure:  # noqa: BLE001 - a crash in the model is a disagreement to count and show
                    cases.append((name, value, 'differs', answer, f'{type(failure).__name__}: {failure}'[:90]))
                    continue
                same = re.sub(r'Book\d+', 'Book', model) == re.sub(r'Book\d+', 'Book', answer)
                cases.append((name, value, 'agrees' if same else 'differs', answer, model))
    return cases


def main() -> None:
    cases = compare()
    totals: dict[str, Counter[str]] = {}
    for name, value, verdict, _, _ in cases:
        totals.setdefault(f"{name.split('.')[0]} {value.split()[0]}", Counter())[verdict] += 1
    for obj, count in totals.items():
        total = sum(count.values())
        measured = total - count['unmeasured']
        print(f"{obj}: {total} cases, {count['agrees']} agree, {count['differs']} differ, {count['reports']} report, "
              f"{count['unmeasured']} unmeasured; disagreement {100 * count['differs'] / max(measured, 1):.1f}%")
    if '--show' in sys.argv:
        for name, value, verdict, answer, model in cases:
            if verdict != 'agrees':
                print(f'{verdict[:6]} {name} = {value}: excel {answer} | model {model}')


if __name__ == '__main__':
    main()
