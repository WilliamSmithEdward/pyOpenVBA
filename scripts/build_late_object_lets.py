"""Write src/pyopenvba/interpreter/_late_object_lets.py from what live Excel answered to an object given by Let.

It reads the two object-Let probes (tests/fixtures/late_object_lets.json, late_object_lets_2.json), the member sweep over Range, Font,
Interior, Worksheet, Workbook and Application (object_let_sweep.json), the Workbook setter held-out rounds
(tests/fixtures/workbook_setters_heldout/round*_excel.json) and the Range-to-Workbook-settings probe (workbook_setters_range.json). For each
member, binding and kind of object a member answers by its own rule (one with no value, a Collection, Nothing, one area of several cells,
several areas), it keeps Excel's error where the set left the member as other refused sets from the same state left it, and 0 where Excel
took the object's value just as the model's value-taking path did, and -1 where Excel took an object with no value, which the member then
receives itself. A kind measured two ways, or one where the set changed the member, is
left out, so the model refuses it openly.

    python scripts/build_late_object_lets.py
"""
from __future__ import annotations

import json
from collections import Counter
from collections.abc import Iterator
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / 'src/pyopenvba/interpreter/_late_object_lets.py'
FIXTURES = ROOT / 'tests/fixtures'
#: The kind each probe's object is, for a member that receives the object.
KINDS = {'ws.Range("A1").Font': 'novalue', 'ws.Range("A1").Interior': 'novalue', 'ws': 'novalue', 'New Collection': 'collection',
         'Nothing': 'nothing', 'ws.Range("A1:C1")': 'cells', 'ws.Range("A1:B1")': 'cells', 'ws.Range("A1,C1")': 'areas'}
#: The kinds that have no value to hand over, so a member that takes one receives the object.
RECEIVED = frozenset({'novalue', 'collection', 'nothing'})
#: A row: binding, class.member, the case's starting state, kind, Excel's answer and the model's before the table, each "eN/read".
Row = tuple[str, str, str, str | None, str, str]


def _probe_rows() -> Iterator[Row]:
    for name in ('late_object_lets', 'late_object_lets_2', 'object_let_sweep'):
        record = json.loads((FIXTURES / f'{name}.json').read_text(encoding='utf-8'))
        before = record['predicted' if 'predicted' in record else 'model']
        for binding in ('typed', 'late'):
            for case, answer, model in zip(record['cases'], record['excel'][binding], before[binding]):
                yield binding, f'{case[0].lower()}.{case[2].lower()}', name, KINDS.get(case[-1]), answer, model


def _setter_rows() -> Iterator[Row]:
    """The setter rounds' cases as rows: the set's error and the setting's own read after it."""
    for path in sorted((FIXTURES / 'workbook_setters_heldout').glob('round*_excel.json')):
        record = json.loads(path.read_text(encoding='utf-8'))
        predicted = json.loads(path.with_name(path.name.replace('_excel', '_predictions')).read_text(encoding='utf-8'))
        names = sorted({case['setting'] for case in record['cases']}, key=[case['setting'] for case in record['cases']].index)
        for declared, binding in (('Workbook', 'typed'), ('Object', 'late')):
            for case, answer, model in zip(record['cases'], record['answers'][declared], predicted['answers'][declared]):
                def own(text: str) -> str:
                    if text.startswith('refused'):
                        return text
                    fields = text.split('|')[0].split('/')
                    return f'{fields[0]}/{fields[1 + names.index(case["setting"])]}'
                yield (binding, f'workbook.{case["setting"].lower()}', f'{path.name}:{case["first"]}', KINDS.get(case['value']),
                       own(answer), own(model))


def _range_rows() -> Iterator[Row]:
    record = json.loads((FIXTURES / 'workbook_setters_range.json').read_text(encoding='utf-8'))
    cases = [(setting, value) for setting in record['settings'] for value in record['ranges']]
    for declared, binding in (('Workbook', 'typed'), ('Object', 'late')):
        for (setting, value), answer, model in zip(cases, record['excel'][declared], record['predicted'][declared]):
            yield binding, f'workbook.{setting.lower()}', 'workbook_setters_range', KINDS.get(value), answer, model


def build() -> dict[str, dict[str, dict[str, int]]]:
    rows = [row for row in (*_probe_rows(), *_setter_rows(), *_range_rows()) if row[4] != 'does not compile']
    # What a member reads when a set changes nothing, by binding, member and starting state: the read most of its refused sets left.
    reads: dict[tuple[str, str, str], Counter[str]] = {}
    for binding, member, state, _kind, answer, _model in rows:
        if not answer.startswith('e0'):
            reads.setdefault((binding, member, state), Counter())[answer.partition('/')[2]] += 1
    unchanged = {key: {counted.most_common(1)[0][0]} for key, counted in reads.items()}
    table: dict[str, dict[str, dict[str, int]]] = {'typed': {}, 'late': {}}
    refused: set[tuple[str, str, str]] = set()
    for binding, member, state, kind, answer, model in rows:
        if kind is None:
            continue
        number, _, read = answer.partition('/')
        code = int(number.removeprefix('e'))
        seen = unchanged.get((binding, member, state), set())
        if not code and kind in RECEIVED:
            # An object with no value that Excel took: the member receives the object itself, and its setter answers it, which the replay
            # of every such case checks (tests/test_vba_object_lets.py).
            code = -1
        if (not code and model == answer) or code == -1 or (code > 0 and len(seen) == 1 and read in seen):
            entry = table[binding].setdefault(member, {})
            if entry.get(kind, code) != code:
                refused.add((binding, member, kind))
            entry[kind] = code
        else:
            refused.add((binding, member, kind))
    for binding, member, kind in refused:
        table[binding].get(member, {}).pop(kind, None)
    return {binding: {member: kinds for member, kinds in sorted(members.items()) if kinds} for binding, members in table.items()}


def main() -> None:
    table = build()
    lines = ['"""What a member answers to an object given by Let, measured per member, as live Excel answered it.', '',
             'Written by scripts/build_late_object_lets.py from tests/fixtures/late_object_lets.json, late_object_lets_2.json,',
             'object_let_sweep.json, the Workbook setter rounds in tests/fixtures/workbook_setters_heldout/ and workbook_setters_range.json.',
             'Do not edit by hand.', '',
             'Keyed by binding, then class.member, then the kind of object: novalue (no default member a Let can read), collection, nothing,',
             'cells (one area of several cells) and areas (several areas). Each is the error Excel gives, the member left as it was, or 0 where',
             "the member takes the object's value, or -1 where it takes an object with no value, which it receives itself. A member or kind",
             'not here is refused openly: the answer goes by the member, not by its type,',
             'so Workbook.Date1904 and Workbook.CheckCompatibility, both Booleans, answer a Font with 1004 and 13.', '"""',
             'from __future__ import annotations', '', 'LATE_OBJECT_LETS: dict[str, dict[str, dict[str, int]]] = {']
    for binding, members in table.items():
        lines.append(f'    {binding!r}: {{')
        lines += [f'        {member!r}: {kinds!r},' for member, kinds in members.items()]
        lines.append('    },')
    lines.append('}')
    OUT.write_text('\n'.join(lines) + '\n', encoding='utf-8', newline='\n')
    print(sum(len(members) for members in table.values()), 'members written')


if __name__ == '__main__':
    main()
