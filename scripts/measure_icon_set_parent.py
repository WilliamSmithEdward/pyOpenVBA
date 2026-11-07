"""What an IconSet's Parent is, by how it was reached. Two rules were written down before Excel answered:

A. Always the workbook's IconSets, as wb.IconSets(1).Parent is (tests/fixtures/object_trio.json).
B. The IconSetCondition it was taken from, when reached through a condition's IconSet; the IconSets when reached through
   wb.IconSets.

Each reach is asked on the kind rounds' three books (new, reopened, the Excel-saved tests/fixtures/kind_rounds/k.xlsx), late
bound and typed, through scripts/measure_kind_rounds.py's setup, and written to tests/fixtures/icon_set_parent.json.

    python scripts/measure_icon_set_parent.py
"""
from __future__ import annotations

import json
import os
import sys
import tempfile
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import measure_kind_rounds as kinds  # noqa: E402

base = kinds.base
OUT = base.FIXTURES / 'icon_set_parent.json'
SEED = 52190
CONDITION = 'ws.Range(rIcon).FormatConditions(1).IconSet'
ASSIGN = 'ws.Range(rIcon).FormatConditions(1).IconSet = wbT.IconSets(5)'
#: (reach, edit before it, call, rule A's answer, rule B's answer)
ASKS = [
    (CONDITION, '', 'x = TypeName(o.Parent)', 'String:IconSets', 'String:IconSetCondition'),
    (CONDITION, ASSIGN, 'x = TypeName(o.Parent)', 'String:IconSets', 'String:IconSetCondition'),
    (CONDITION, '', 'x = TypeName(o.Parent.Parent)', 'String:Workbook', 'String:Range'),
    (CONDITION, '', 'x = Key(o.Parent.Parent, book)', 'Workbook:<book>', 'Range:' + 'rIcon'),
    ('wbT.IconSets(3)', '', 'x = TypeName(o.Parent)', 'String:IconSets', 'String:IconSets'),
    ('wbT.IconSets(3)', '', 'x = TypeName(o.Parent.Parent)', 'String:Workbook', 'String:Workbook'),
    ('wbT.IconSets', '', 'x = TypeName(o.Item(2).Parent)', 'String:IconSets', 'String:IconSets'),
]


def cases() -> list[dict[str, str]]:
    out = []
    for kind in base.KINDS:
        for reach, edit, call, rule_a, rule_b in ASKS:
            cls = 'IconSets' if reach == 'wbT.IconSets' else 'IconSet'
            out.append({'cls': cls, 'reach': reach, 'edit': edit, 'held': 'False', 'kind': kind, 'call': call.replace('x = TypeName', 'x = "String:" & TypeName'),
                        'rule_a': rule_a, 'rule_b': rule_b})
    return out


def main() -> None:
    from pyvbaharness import ExcelSession, HarnessConfig

    asked = cases()
    for case in asked:
        case['model_late'] = base.model_answer(case, False, SEED, kinds.KINDS)
        case['model_typed'] = base.model_answer(case, True, SEED, kinds.KINDS)
    lock = base._lock()
    while True:
        try:
            os.mkdir(lock)
            break
        except FileExistsError:
            time.sleep(10)
    try:
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as held, ExcelSession(HarnessConfig(lock_wait_s=1200.0, exclusive=False)) as excel:
            folder = Path(held)
            kinds.books(folder)
            excel.new_document()
            for case in asked:
                for tag, typed in (('late', False), ('typed', True)):
                    done = excel.run_vba(base.source(case, typed, SEED, kinds.KINDS), 'One', (str(folder),), timeout=60.0, line_numbers=False)
                    case[tag] = str(done.value) if done.ok else f'no answer: {str(done)[:300]}'
    finally:
        os.rmdir(lock)
    OUT.write_text(json.dumps({'seed': SEED, 'cases': asked}, indent=1) + '\n', encoding='utf-8')
    for case in asked:
        print(case['kind'], case['reach'][:30], case['edit'][:12], case['call'][:40], '| Excel', case['late'], '/', case['typed'], '| A', case['rule_a'], '| B', case['rule_b'],
              '| model', case['model_late'], '/', case['model_typed'])


if __name__ == '__main__':
    main()
